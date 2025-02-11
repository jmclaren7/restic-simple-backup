#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=include\SimpleBackup.ico
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Comment=https://github.com/jmclaren7/restic-simple-backup
#AutoIt3Wrapper_Res_Description=SimpleBackup
#AutoIt3Wrapper_Res_Fileversion=1.0.0.280
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_ProductVersion=1
#AutoIt3Wrapper_Res_LegalCopyright=SimpleBackup
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_requestedExecutionLevel=highestAvailable
#AutoIt3Wrapper_Run_Au3Stripper=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

;#RequireAdmin

#include <Array.au3>
#include <AutoItConstants.au3>
#include <Crypt.au3>
#include <File.au3>
#include <GuiComboBox.au3>
#include <GUIConstantsEx.au3>
#include <GuiEdit.au3>
#include <GuiMenu.au3>
#include <StaticConstants.au3>
#include <String.au3>
#include <WinAPIConv.au3>
#include <WinAPIDiag.au3>
#include <WinAPISysWin.au3>
#include <WindowsConstants.au3>

; https://github.com/jmclaren7/autoit-scripts/blob/master/CommonFunctions.au3
#include <include\CommonFunctions.au3>
#include <include\Console.au3>

; Setup version and program title globals
Global $Version = 0
If @Compiled Then $Version = FileGetVersion(@AutoItExe)
Global $Title = StringTrimRight(@ScriptName, 4)
Global $TitleVersion = $Title & " v" & StringTrimLeft($Version, StringInStr($Version, ".", 0, -1))

; Setup Logging
_Console_Attach() ; If it was launched from a console, attach to that console
Global $LogFileMaxSize = 512
Global $LogLevel = 1
If @Compiled Then
	$LogFullPath = StringTrimRight(@ScriptFullPath, 4) & ".log"
	_Console_Alloc()
Else
	$LogLevel = 3
	Global $LogTitle = "Log - " & $Title
	Global $LogWindowStart = 1
EndIf

_Log("Starting " & $TitleVersion)

; Setup some miscellaneous globals
Global $TempDir = _TempFile(@TempDir, "sbr", "tmp", 10)
Global $ResticFullPath = $TempDir & "\restic.exe"
Global $ResticBrowserFullPath = $TempDir & "\Restic-Browser.exe"
Global $SMTPSettings = "Backup_Name|SMTP_Server|SMTP_UserName|SMTP_Password|SMTP_FromAddress|SMTP_FromName|SMTP_ToAddress|SMTP_SendOnFailure|SMTP_SendOnSuccess"
Global $WebhookSettings = "WebhookURL_Success|WebhookURL_Failure"
Global $InternalSettings = "Setup_Password|Backup_Path|Backup_Prune|" & $SMTPSettings
Global $RequiredSettings = "Setup_Password|Backup_Path|Backup_Prune|RESTIC_REPOSITORY|RESTIC_PASSWORD"
Global $SettingsTemplate = $RequiredSettings & "|AWS_ACCESS_KEY_ID|AWS_SECRET_ACCESS_KEY|RESTIC_READ_CONCURRENCY=4|RESTIC_PACK_SIZE=32|" & $SMTPSettings & "|" & $WebhookSettings
Global $ActiveConfigFileFullPath = _GetProfileFullPath()

; $RunSTDIO will determine how we execute restic, default $STDERR_MERGED will allow restic output to be logged to file
Global $RunSTDIO = $STDERR_MERGED

; These SHA1 hashes are used to verify the Restic and Restic-Browser binaries right before they run
Global $SafeHash = _
		"0x0096DB4992E253BF4D59CC8B431046FA227B2B56" & _ ; 01/25/25
		"0x389f226485ac7c1a4987380971e0ee19623339fb" & _ ; 12/26/24
		"0x1501b645438de5744d39e87204206a9faa8df62a" & _ ; 11/11/24
		"0xfc731979ce12a857efd2f210f51dcd5d08f66d24" & _ ; 11/01/24
		"0xc3cf4db47ffe72924bda7bc47231dda270b517a2"     ; 11/01/24

; Register our exit function for cleanup
OnAutoItExitRegister("_Exit")

; This key is used to encrypt the configuration file but is mostly just going to limit non-targeted/low-effort attacks, customizing the key for your own deployment could help though
Global $Key
$Key = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\RSB", "k")
If @error Or $Key = "" Then
	$Key = _WinAPI_UniqueHardwareID($UHID_MB) & DriveGetSerial(@HomeDrive & "\") & @CPUArch
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\RSB", "k", "REG_SZ", $Key)
	If @error Then
		_Log("Error saving key")
	Else
		_Log("Saved key")
	EndIf
Else
	_Log("Key loaded from registry", 2)

EndIf

; Interpret command line parameters
$Command = Default
For $i = 1 To $CmdLine[0]
	_Log("Parameter: " & $CmdLine[$i])
	Switch $CmdLine[$i]
		Case "profile"
			; The profile parameter and the following parameter are used to adjust the
			$i = $i + 1
			$ActiveConfigFileFullPath = _GetProfileFullPath($CmdLine[$i])
			_Log("Config file is now " & $ActiveConfigFileFullPath)

		Case Else
			; The first parameter not specialy handled must be the command
			If $Command = Default Then $Command = $CmdLine[$i]

	EndSwitch
Next

; Default command for when no parameters are provided
If $Command = Default Then $Command = "setup"
_Log("Command: " & $Command)

; This loop is used to effectively restart the program for when we switch profiles
While 1
	; Load config from file and load to array
	$ConfigData = _ReadConfig()
	If @error = 1 Then MsgBox(16, $Title, "Error opening configuration")
	Global $aConfig = _ConfigToArray($ConfigData)

	; If config is empty load it with the template, otherwise make sure it at least has required settings
	If UBound($aConfig) = 0 Then
		_ForceRequiredConfig($aConfig, $SettingsTemplate)
	Else
		_ForceRequiredConfig($aConfig, $RequiredSettings)
	EndIf

	; Continue based on command line or default command
	Switch $Command
		; Display help information
		Case "help", "/?"
			_Log("Valid Restic Commands: version, stats, init, check, snapshots, backup, --help")
			_Log("Valid Script Commands: setup, command")

		; Basic commands allowed to be passed to the Restic executable
		Case "version", "stats", "init", "check", "snapshots", "--help"
			_Restic($Command)

		; Pass arbitrary commands to the Restic executable
		Case "c", "command"
			_Auth()

			$Run = StringTrimLeft($CmdLineRaw, StringLen($CmdLine[1]) + 1)
			_Restic($Run)

		; Backup command
		Case "backup"
			If _KeyValue($aConfig, "Backup_Path") = "" Then
				_Log("Backup path not set, skipping")
				ContinueCase
			EndIf

			$Result = _Restic("backup " & _KeyValue($aConfig, "Backup_Path") & " --no-scan")
			$BackupSuccess = StringRegExp($Result, "snapshot [0-9a-fA-F]+ saved")

			; If the backup result and email options match, send an email
			If (Not $BackupSuccess And _KeyValue($aConfig, "SMTP_SendOnFailure")) Or ($BackupSuccess And _KeyValue($aConfig, "SMTP_SendOnSuccess")) Then
				If $BackupSuccess Then
					$sSubject = "[Completed]"
				Else
					$sSubject = "[Failed]"
				EndIf
				$sSubject = _KeyValue($aConfig, "Backup_Name") & " " & $sSubject & " - " & $Title
				$sSubject = StringStripWS($sSubject, 1)
				$sBody = $Result

				_Log("Sending Email " & $sSubject)
				_INetSmtpMailCom(_KeyValue($aConfig, "SMTP_Server"), _KeyValue($aConfig, "SMTP_FromName"), _KeyValue($aConfig, "SMTP_FromAddress"), _
						_KeyValue($aConfig, "SMTP_ToAddress"), $sSubject, $sBody, _KeyValue($aConfig, "SMTP_UserName"), _KeyValue($aConfig, "SMTP_Password"))

			EndIf

			; Trigger webhook based on success or failure
			If Not $BackupSuccess And _KeyValue($aConfig, "WebhookURL_Failure") Then
				_Log("Failure Inet: " & BinaryToString(InetRead(_KeyValue($aConfig, "WebhookURL_Failure"), 1)))

			EndIf

			If $BackupSuccess And _KeyValue($aConfig, "WebhookURL_Success") Then
				_Log("Success Inet: " & BinaryToString(InetRead(_KeyValue($aConfig, "WebhookURL_Success"), 1)))
			EndIf

			$Prune = _KeyValue($aConfig, "Backup_Prune")
			If $Prune Then _Restic("forget --prune " & $Prune)

		; Setup GUI
		Case "setup"
			WinMove("[TITLE:" & @AutoItExe & "; CLASS:ConsoleWindowClass]", "", 4, 4)

			_Auth()

			Global $SettingsForm, $RunCombo

			#Region ### START Koda GUI section ###
			$SettingsForm = GUICreate("Title", 601, 533, -1, -1, BitOR($GUI_SS_DEFAULT_GUI, $WS_SIZEBOX, $WS_THICKFRAME))
			$ApplyButton = GUICtrlCreateButton("Apply", 510, 496, 75, 25)
			GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
			$ScriptEdit = GUICtrlCreateEdit("", 7, 3, 585, 441, BitOR($GUI_SS_DEFAULT_EDIT, $WS_BORDER), 0)
			GUICtrlSetData(-1, "")
			GUICtrlSetFont(-1, 10, 400, 0, "Consolas")
			GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKBOTTOM)
			$CancelButton = GUICtrlCreateButton("Cancel", 425, 496, 75, 25)
			GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
			$OKButton = GUICtrlCreateButton("OK", 339, 496, 75, 25)
			GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
			$RunButton = GUICtrlCreateButton("Run", 532, 456, 51, 33)
			GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
			$RunCombo = GUICtrlCreateCombo("Select or Type A Command", 15, 462, 505, 25, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL))
			GUICtrlSetFont(-1, 9, 400, 0, "Consolas")
			GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKHEIGHT)
			GUISetState(@SW_SHOW)
			#EndRegion ### END Koda GUI section ###

			; Set some of the GUI parameters that we don't or can't do in Koda
			WinSetTitle($SettingsForm, "", $TitleVersion) ; Set the title from title variable
			If _GetProfileName() <> "Default" Then WinSetTitle($SettingsForm, "", $TitleVersion & " - " & _GetProfileName()) ; Change title to include profile if not the default profile
			GUICtrlSetData($ScriptEdit, _ArrayToConfig($aConfig)) ; Load the edit box with config data
			_GUICtrlComboBox_SetDroppedWidth($RunCombo, 600) ; Set the width of the combobox drop down beyond the width of the combobox
			_UpdateCommandComboBox() ; Set the options in the combobox

			; Setup a custom menu
			Global $MenuMsg = 0
			If Not IsDeclared("ExitMenuItem") Then Global Enum $ExitMenuItem = 1000, $ScheduledTaskMenuItem, $FixConsoleMenuItem, $BrowserMenuItem, _
					$VerboseMenuItem, $TemplateMenuItem, $NewProfileMenuItem, $AboutMenuItem, $WebsiteMenuItem, $InstallMenuItem

			; Setup file menu
			$g_hFile = _GUICtrlMenu_CreateMenu()
			; Create menu items based on config files found
			Global $aConfigFiles = _FileListToArray(@ScriptDir, StringTrimRight(@ScriptName, 4) & "*.dat", 1, True)
			For $i = 1 To UBound($aConfigFiles) - 1
				$ThisProfileName = _GetProfileName($aConfigFiles[$i])
				_Log("Found profile: " & $ThisProfileName)

				$cmdID = 1100 + $i

				; If this menu item is for the active config then add text to let user know
				If $aConfigFiles[$i] = $ActiveConfigFileFullPath Then $ThisProfileName = $ThisProfileName & " (Current Profile)"

				; Create the menu item for the profile
				_GUICtrlMenu_InsertMenuItem($g_hFile, -1, $ThisProfileName, $cmdID)

				; If this menu item is for the active config then disable clicking on it
				If $aConfigFiles[$i] = $ActiveConfigFileFullPath Then _GUICtrlMenu_SetItemDisabled($g_hFile, $cmdID, True, False)
			Next
			_GUICtrlMenu_InsertMenuItem($g_hFile, -1, "New Profile/Config...", $NewProfileMenuItem)
			_GUICtrlMenu_InsertMenuItem($g_hFile, -1, "")
			_GUICtrlMenu_InsertMenuItem($g_hFile, -1, "Exit", $ExitMenuItem)

			; Setup tools menu
			$g_hTools = _GUICtrlMenu_CreateMenu()
			_GUICtrlMenu_InsertMenuItem($g_hTools, -1, "Copy " & @ScriptName & " To Program Files And Restart Program", $InstallMenuItem)
			_GUICtrlMenu_InsertMenuItem($g_hTools, -1, "Create/Reset Scheduled Task (Profile: " & _GetProfileName() & ")", $ScheduledTaskMenuItem)
			_GUICtrlMenu_InsertMenuItem($g_hTools, -1, "Open Restic Browser", $BrowserMenuItem)
			_GUICtrlMenu_InsertMenuItem($g_hTools, -1, "Add Missing Configuration Options From Template", $TemplateMenuItem)

			; Setup advanced menu
			$g_hAdvanced = _GUICtrlMenu_CreateMenu()
			_GUICtrlMenu_InsertMenuItem($g_hAdvanced, -1, "Make Restic Progress Bar Work In Console (breaks file log)", $FixConsoleMenuItem)
			_GUICtrlMenu_InsertMenuItem($g_hAdvanced, -1, "Verbose Logs (While in GUI, may contain secrets)", $VerboseMenuItem)

			; Setup help menu
			$g_hHelp = _GUICtrlMenu_CreateMenu()
			_GUICtrlMenu_InsertMenuItem($g_hHelp, -1, "Visit Website", $WebsiteMenuItem)
			_GUICtrlMenu_InsertMenuItem($g_hHelp, -1, "About " & $Title, $AboutMenuItem)

			; Setup main menu
			$g_hMain = _GUICtrlMenu_CreateMenu(BitOR($MNS_CHECKORBMP, $MNS_MODELESS))
			_GUICtrlMenu_InsertMenuItem($g_hMain, -1, "&File", 0, $g_hFile)
			_GUICtrlMenu_InsertMenuItem($g_hMain, -1, "&Tools", 0, $g_hTools)
			_GUICtrlMenu_InsertMenuItem($g_hMain, -1, "&Advanced", 0, $g_hAdvanced)
			_GUICtrlMenu_InsertMenuItem($g_hMain, -1, "&Help", 0, $g_hHelp)

			; Create menu
			_GUICtrlMenu_SetMenu($SettingsForm, $g_hMain)

			; Additonal menu setup
			If @Compiled Then _GUICtrlMenu_SetItemState($g_hMain, $FixConsoleMenuItem, $MFS_CHECKED, True, False)
			If $LogLevel = 3 Then _GUICtrlMenu_SetItemState($g_hMain, $VerboseMenuItem, $MFS_CHECKED, True, False)
			GUIRegisterMsg($WM_COMMAND, "_WM_COMMAND")


			; Setup GUI accelerators
			Dim $aAccelKeys[2][2]

			; CTRL+S will work as apply button
			$aAccelKeys[0][0] = "^s"
			$aAccelKeys[0][1] = $ApplyButton

			; ENTER will run the command while combobox has focus
			$cEnterPressed = GUICtrlCreateDummy()
			$aAccelKeys[1][0] = "{ENTER}"
			$aAccelKeys[1][1] = $cEnterPressed

			GUISetAccelerators($aAccelKeys)

			; GUI loop
			While 1
				$nMsg = GUIGetMsg()
				; Add menu actions from custom menu gui
				If $nMsg = 0 And $MenuMsg <> 0 Then
					$nMsg = $MenuMsg
					$MenuMsg = 0
				EndIf
				If $nMsg <> 0 And Not ($nMsg > -13 And $nMsg < -3) Then _Log("Merged $nMsg = " & $nMsg, 3)

				; Continue based on GUI action
				Switch $nMsg
					; Save or save and close
					Case $ApplyButton, $OKButton
						_Log("Apply/Ok")
						$aConfig = _ConfigToArray(GUICtrlRead($ScriptEdit))
						_WriteConfig(_ArrayToConfig($aConfig))

						; Warn user if the backup path doesn't make sense
						If _KeyValue($aConfig, "Backup_Path") <> "" Then
							$ErrorMessage = "Back_Path might contain an invalid path, your settings have been saved but please verify you have specified a valid path and used quotes properly. This warning was triggered based on the follow text..."
							$sBackup_Path = _KeyValue($aConfig, "Backup_Path")
							$aBackup_Path = StringRegExp($sBackup_Path, '["''].*?["'']', $STR_REGEXPARRAYGLOBALMATCH)

							; Check for paths inside quoted strings
							If IsArray($aBackup_Path) Then
								For $i = 0 To UBound($aBackup_Path) - 1
									$StrippedPath = StringReplace($aBackup_Path[$i], """", "")
									If Not FileExists($StrippedPath) Then MsgBox($MB_ICONINFORMATION, $TitleVersion, $ErrorMessage & @CRLF & @CRLF & $aBackup_Path[$i])

								Next
							; No quoted strings were found so check the entire string
							Else
								$StrippedPath = StringReplace($sBackup_Path, """", "") ; todo: only trim quotes from first and last
								If Not FileExists($StrippedPath) Then
									; The entire string wasn't a path so test up to the first space
									; This wont capture issues if we have more than one path seperated by space but that would be hard to do if we also want to access parameters here
									$StrippedPath = StringLeft($StrippedPath, StringInStr($StrippedPath, " "))
									If Not FileExists($StrippedPath) Then MsgBox($MB_ICONINFORMATION, $TitleVersion, $ErrorMessage & @CRLF & @CRLF & $sBackup_Path)
								EndIf
							EndIf
						EndIf
						If $nMsg = $OKButton Then Exit

						; Since we just changed some settings, update the combo box which might be using data from our settings
						_UpdateCommandComboBox()

					; Close program
					Case $GUI_EVENT_CLOSE, $CancelButton, $ExitMenuItem
						; Exit and run the registered exit function for cleanup
						Exit

					; Custom menu items that use checkboxes need to be check and unchecked manually when clicked
					Case $FixConsoleMenuItem, $VerboseMenuItem
						If _GUICtrlMenu_GetItemChecked($g_hMain, $nMsg, False) Then
							_GUICtrlMenu_SetItemChecked($g_hMain, $nMsg, False, False)
						Else
							_GUICtrlMenu_SetItemChecked($g_hMain, $nMsg, True, False)
						EndIf

					; Add missing key=value pairs to existing config
					Case $TemplateMenuItem
						_ForceRequiredConfig($aConfig, $SettingsTemplate)
						GUICtrlSetData($ScriptEdit, _ArrayToConfig($aConfig)) ; Load the edit box with config data

					; Open the github page
					Case $WebsiteMenuItem
						ShellExecute("https://github.com/jmclaren7/restic-simple-backup")

					; Open a dialog with program information
					Case $AboutMenuItem
						MsgBox($MB_ICONINFORMATION, $TitleVersion, _
								"Restic SimpleBackup" & @CRLF & "https://github.com/jmclaren7/restic-simple-backup" & @CRLF & "Copyright (c) 2023, John McLaren" & @CRLF & @CRLF & _
								"Restic" & @CRLF & "https://github.com/restic/restic" & @CRLF & "Copyright (c) 2014, Alexander Neumann" & @CRLF & @CRLF & _
								"Restic Browser" & @CRLF & "https://github.com/emuell/restic-browser" & @CRLF & "Copyright (c) 2022, Eduard Müller / taktik", 0, $SettingsForm)

					; Create or switch profile
					Case $NewProfileMenuItem, 1100 To 1199
						_Log("Profile create/switch")

						If $nMsg = $NewProfileMenuItem Then
							; Prompt for profile name, if the profile name already exists it will be loaded
							$NewProfile = InputBox($TitleVersion, "Enter a name for the new profile/config", "", "", Default, 130)
							If @error Then ContinueLoop
							$ActiveConfigFileFullPath = _GetProfileFullPath($NewProfile)
						Else
							$ActiveConfigFileFullPath = $aConfigFiles[$nMsg - 1100]

						EndIf

						_Log("$ActiveConfigFileFullPath=" & $ActiveConfigFileFullPath, 3)

						; Delete the GUI and restart
						GUIDelete($SettingsForm)
						ContinueLoop 2

					; Start the Restic-Browser
					Case $BrowserMenuItem
						; Install Restic-Browser.exe
						; If the executable exists, verify it
						If FileExists($ResticBrowserFullPath) Then
							Local $Hash = _Crypt_HashFile($ResticBrowserFullPath, $CALG_SHA1)
							If Not StringInStr($SafeHash, $Hash) Then
								_Error("Error starting Restic-Browser.exe, the program will now exit", Default, "Hash error - " & $Hash)
								Exit
							EndIf
						Else
							; Pack and unpack the executable
							DirCreate($TempDir)
							If FileInstall("include\Restic-Browser.exe", $ResticBrowserFullPath, 1) = 0 Then
								_Error("Error starting Restic-Browser.exe, the program will now exit", Default, "FileInstall error - " & @error)
								Exit
							EndIf
							; Verify the hash of the executable
							Local $Hash = _Crypt_HashFile($ResticBrowserFullPath, $CALG_SHA1)
							If Not StringInStr($SafeHash, $Hash) Then
								_Error("Error starting Restic-Browser.exe, the program will now exit", Default, "Hash error - " & $Hash)
								Exit
							EndIf
						EndIf

						; Run a dummy restic command to unpack the restic executable
						_Restic("version")

						; Update PATH env so that Restic-browser.exe can start restic.exe
						$EnvPath = EnvGet("Path")
						If Not StringInStr($EnvPath, $TempDir) Then
							EnvSet("Path", $TempDir & ";" & $EnvPath)
							_Log("EnvSet: " & @error, 3)
						EndIf

						; Verify the hash of the restic-browser.exe
						Local $Hash = _Crypt_HashFile($ResticBrowserFullPath, $CALG_SHA1)
						If Not StringInStr($SafeHash, $Hash) Then
							_Log("Hash error - " & $Hash)
							MsgBox(16, $Title, "Error starting program")
							Exit
						EndIf

						; Load the Restic credential envs and start restic-browser.exe
						_UpdateEnv($aConfig)
						$ResticBrowserPid = Run($ResticBrowserFullPath)

					; Copy the program to Program Files
					Case $InstallMenuItem
						If @Compiled Then
							$sDestinationFullPath = @ProgramFilesDir & "\" & StringTrimRight(@ScriptName, 4) & "\" & @ScriptName
							FileCopy(@ScriptFullPath, $sDestinationFullPath, $FC_CREATEPATH)
							ShellExecute($sDestinationFullPath)
							Exit
						EndIf

					; Create a sceduled task to run the backup
					Case $ScheduledTaskMenuItem
						If _GetProfileName() <> "Default" Then
							$ProfileSwitch = " profile " & _GetProfileName()
							$TaskName = "." & _GetProfileName()
						Else
							$ProfileSwitch = ""
							$TaskName = ""
						EndIf

						$Run = "SCHTASKS /CREATE /SC DAILY /TN " & $Title & $TaskName & " /TR ""'" & @ScriptFullPath & "' backup" & $ProfileSwitch & """ /ST 22:00 /RL Highest /NP /F /RU System"
						_Log($Run)
						$Return = _RunWait($Run, @ScriptDir, @SW_SHOW, $STDERR_MERGED, True)
						If StringInStr($Return, "SUCCESS: ") Then
							MsgBox(0, $TitleVersion, "Scheduled task created. Please review and test the task.")
						Else
							MsgBox($MB_ICONERROR, $TitleVersion, "Error creating scheduled task.")
							If Not @Compiled Then MsgBox($MB_ICONERROR, $TitleVersion, "Are you running without admin?")
						EndIf

					; Run the command provided from the combobox
					Case $RunButton, $cEnterPressed
						If $nMsg = $cEnterPressed And ControlGetFocus($SettingsForm) <> "Edit2" Then ContinueLoop

						; Adjust window visibility and activation due to long running process
						GUISetState(@SW_DISABLE, $SettingsForm)
						WinSetTrans($SettingsForm, "", 210)
						If @Compiled Then WinActivate("[TITLE:" & @AutoItExe & "; CLASS:ConsoleWindowClass]")

						; Don't continue if combo box is on the placeholder text
						If _GUICtrlComboBox_GetCurSel($RunCombo) = 0 Then ContinueLoop

						; Continue based on combobox value
						$RunComboText = GUICtrlRead($RunCombo)
						Switch $RunComboText
							; Custom/advanced commands can be added here in the future
							Case "cmd"
								EnvSet("Path", EnvGet("Path") & ";" & $TempDir)
								_Restic("stage env")
								Run("cmd.exe")

							Case "Test Email"
								If _KeyValue($aConfig, "SMTP_Server") Then
									$sSubject = "Test Subject"
									$sBody = "Test Body"
									$Return = _INetSmtpMailCom(_KeyValue($aConfig, "SMTP_Server"), _KeyValue($aConfig, "SMTP_FromName"), _KeyValue($aConfig, "SMTP_FromAddress"), _KeyValue($aConfig, "SMTP_ToAddress"), _
											$sSubject, $sBody, _KeyValue($aConfig, "SMTP_UserName"), _KeyValue($aConfig, "SMTP_Password"))
									_Log("Email Test: $Return=" & $Return & "  @error=" & @error)
								EndIf

							Case Else
								$PID = _Restic($RunComboText)

						EndSwitch
				EndSwitch

				; Adjust $LogLevel
				If _GUICtrlMenu_GetItemChecked($g_hMain, $VerboseMenuItem, False) Then
					$LogLevel = 3
				Else
					$LogLevel = 1
				EndIf

				; Adjust the global used to determine STDIO streams for child processes
				If _GUICtrlMenu_GetItemChecked($g_hMain, $FixConsoleMenuItem, False) Then
					$RunSTDIO = 0
				Else
					$RunSTDIO = $STDERR_MERGED
				EndIf

				; Re-enable the window when the script resumes from a blocking task
				If BitAND(WinGetState($SettingsForm), @SW_DISABLE) Then
					GUISetState(@SW_ENABLE, $SettingsForm)
					WinSetTrans($SettingsForm, "", 255)
				EndIf

				; Removes help text from combobox selection
				$RunComboText = GUICtrlRead($RunCombo)
				If StringRight($RunComboText, 1) = ")" And StringInStr($RunComboText, "  (") And Not _GUICtrlComboBox_GetDroppedState($RunCombo) Then
					$NewText = StringLeft($RunComboText, StringInStr($RunComboText, "  (") - 1)
					_GUICtrlComboBox_SetEditText($RunCombo, $NewText)
				EndIf

				Sleep(10)
			WEnd

		Case Else
			_Log("Invalid Command")

	EndSwitch

	Exit ; To support existing program flow since adding loop used to restart GUI
WEnd

;=====================================================================================
;=====================================================================================
; Extract profile name from profile path
Func _GetProfileName($sPath = Default)
	If $sPath = Default Then $sPath = $ActiveConfigFileFullPath

	Local $Return = StringRegExp($sPath, "\" & $Title & ".([0-9a-zA-Z.\-_ ]+).dat", 1)

	If @error Then
		Return "Default"
	Else
		Return $Return[0]
	EndIf
EndFunc   ;==>_GetProfileName

; Construct profile path from profile name
Func _GetProfileFullPath($sProfile = Default)
	If $sProfile = Default Then
		$sProfile = ""
	Else
		$sProfile = "." & $sProfile
	EndIf

	Return StringTrimRight(@ScriptFullPath, 4) & $sProfile & ".dat"

EndFunc   ;==>_GetProfileFullPath

; Special function to handle messages from custom GUI menu
Func _WM_COMMAND($hWnd, $iMsg, $wParam, $lParam)
	Local $Temp = _WinAPI_LoWord($wParam)

	;_Log("_WM_COMMAND ($wParam = " & $Temp & ") ", 3)

	If $Temp >= 1000 And $Temp < 1999 Then
		_Log("Updated $MenuMsg To " & $Temp, 3)
		Global $MenuMsg = $Temp
	EndIf

	Return $GUI_RUNDEFMSG
EndFunc   ;==>_WM_COMMAND

; Prompt for a password before continuing
Func _Auth()
	_Log("_Auth", 3)
	Local $InputPass

	While 1
		; Check input first to deal with empty password
		If $InputPass = _KeyValue($aConfig, "Setup_Password") Then ExitLoop

		$InputPass = InputBox($TitleVersion, "Enter Password", "", "*", Default, 130)
		If @error Then Exit

	WEnd

	Return
EndFunc   ;==>_Auth

; Updates the options in the GUI combo box
Func _UpdateCommandComboBox()
	_Log("_UpdateCommandComboBox", 3)

	_GUICtrlComboBox_ResetContent($RunCombo)

	$Opts = "Select or type a Restic command"
	$Opts &= "|" & "init  (Create the Restic repository)"
	$Opts &= "|" & "backup " & _KeyValue($aConfig, "Backup_Path") & "  (Runs a backup)"
	$Opts &= "|" & "forget --prune " & _KeyValue($aConfig, "Backup_Prune") & "  (Removes old backups)"
	$Opts &= "|" & "snapshots  (Lists snapshots in the repository)"
	$Opts &= "|" & "unlock  (Unlocks the repository in case Restic had an issue)"
	$Opts &= "|" & "check --read-data  (Verifies all data in repo SLOW!!!)"
	$Opts &= "|" & "stats raw-data  (Show storage used)"
	$Opts &= "|" & "version  (Show Restic version information)"
	$Opts &= "|" & "--help  (Show Restic help information)"
	GUICtrlSetData($RunCombo, $Opts, "Select or type a Restic command")

EndFunc   ;==>_UpdateCommandComboBox

; Update the environmental variables
Func _UpdateEnv($aArray, $Delete = Default)
	_Log("_UpdateEnv", 3)

	If $Delete = Default Then $Delete = False

	Local $aInternalSettings = StringSplit($InternalSettings, "|")

	; Loop array and set or delete Env
	For $i = 1 To $aArray[0][0]
		; Skip If the value is a comment, empty or internal
		If StringLeft($aArray[$i][0], 1) = "#" Then ContinueLoop
		If $aArray[$i][0] = "" Or $aArray[$i][1] = "" Then ContinueLoop
		If _ArraySearch($aInternalSettings, $aArray[$i][0], 0, 0, 0, 0) <> -1 Then ContinueLoop

		If $Delete Then
			_Log("  EnvSet (Delete): " & $aArray[$i][0], 3)
			EnvSet($aArray[$i][0], "")
		Else
			_Log("  EnvSet: " & $aArray[$i][0], 3)
			EnvSet($aArray[$i][0], $aArray[$i][1])
		EndIf
	Next

	Return
EndFunc   ;==>_UpdateEnv

; Force required keys to show in the config array
Func _ForceRequiredConfig(ByRef $aArray, $sRequired)
	Local $aKeyValuem, $sValue

	$sRequired = StringSplit($sRequired, "|")

	For $i = 1 To $sRequired[0]
		$aKeyValue = StringSplit($sRequired[$i], "=")
		_KeyValue($aArray, $aKeyValue[1])

		If $aKeyValue[0] = 2 Then
			$sValue = $aKeyValue[2]
		Else
			$sValue = ""
		EndIf

		If @error Then _KeyValue($aArray, $aKeyValue[1], $sValue)
	Next
EndFunc   ;==>_ForceRequiredConfig

; Converts string of key=value pairs to array and handles comments
Func _ConfigToArray($ConfigData)
	_Log("_ConfigToArray", 3)

	Local $aConfigLines = StringSplit($ConfigData, @CRLF, 1)
	Local $ThisKey, $KeyValue, $aArray[]

	; Loop through each line
	For $i = 1 To $aConfigLines[0]
		; Prefix comments with #<UID>= so that we can treat them as a key=value pair, each with a unique key
		If StringLeft($aConfigLines[$i], 1) = "#" Then $aConfigLines[$i] = "#" & $i & "=" & $aConfigLines[$i]

		$ThisKey = StringLeft($aConfigLines[$i], StringInStr($aConfigLines[$i], "=") - 1)
		$ThisKey = StringStripWS($ThisKey, 1 + 2)

		If $ThisKey = "" Then ContinueLoop

		$KeyValue = StringTrimLeft($aConfigLines[$i], StringInStr($aConfigLines[$i], "="))
		$KeyValue = StringStripWS($KeyValue, 1 + 2)

		_KeyValue($aArray, $ThisKey, $KeyValue)
	Next


	Return $aArray
EndFunc   ;==>_ConfigToArray

; Converts array back to a string of key=value pairs and fix comments
Func _ArrayToConfig($aArray)
	_Log("_ArrayToConfig", 3)
	Local $ConfigData, $Add

	If UBound($aArray, 0) <> 2 Then Return SetError(1, 0, "")

	; Loop array and combine key=value pairs
	For $i = 1 To $aArray[0][0]
		; If the value is a comment remove the placeholder key
		If StringLeft($aArray[$i][0], 1) = "#" Then
			$Add = $aArray[$i][1]
		ElseIf $aArray[$i][0] = "" Then
			ContinueLoop
		Else
			$Add = $aArray[$i][0] & "=" & $aArray[$i][1]
		EndIf

		$ConfigData = $ConfigData & $Add & @CRLF
	Next

	Return $ConfigData
EndFunc   ;==>_ArrayToConfig

; Read and decrypt the config file
Func _ReadConfig()
	_Log("_ReadConfig", 3)
	Global $Key, $ActiveConfigFileFullPath

	If Not FileExists($ActiveConfigFileFullPath) Then
		_Log("Selected config file doesn't exist")
		Return SetError(2, 0)
	EndIf

	Local $ConfigData = FileRead($ActiveConfigFileFullPath)

	; Decypt Data
	$ConfigData = _Crypt_DecryptData($ConfigData, $Key, $CALG_AES_256)
	If @error OR $ConfigData = "" Then
		_Log("Could not decrypt configuration file")
		_Log("Key: " & $Key, 2)
		Return SetError(1, 0, $ConfigData)
	EndIf

	$ConfigData = BinaryToString($ConfigData)

	Return $ConfigData
EndFunc   ;==>_ReadConfig

; Encrypt and write the config file
Func _WriteConfig($ConfigData)
	_Log("_WriteConfig", 3)
	Global $Key

	; Encrypt
	$ConfigData = _Crypt_EncryptData($ConfigData, $Key, $CALG_AES_256)

	$hConfigFile = FileOpen($ActiveConfigFileFullPath, 2)
	FileWrite($hConfigFile, $ConfigData)

	Return
EndFunc   ;==>_WriteConfig

; Execute a Restic command
Func _Restic($Command)
	_Log("_Restic", 3)

	Global $RunSTDIO, $TempDir, $ResticFullPath

	; Install restic.exe
	; If the executable exists, verify it
	If FileExists($ResticFullPath) Then
		Local $Hash = _Crypt_HashFile($ResticFullPath, $CALG_SHA1)
		If Not StringInStr($SafeHash, $Hash) Then
			_Error("Error starting restic.exe", Default, "Hash error - " & $Hash)
			Exit
		EndIf
	Else
		; Pack/unpack the executable
		DirCreate($TempDir)
		If FileInstall("include\restic.exe", $ResticFullPath, 1) = 0 Then
			_Error("Error starting restic.exe", Default, "FileInstall error - " & @error)
			Exit
		EndIf
		; Verify the hash of the executable
		Local $Hash = _Crypt_HashFile($ResticFullPath, $CALG_SHA1)
		If Not StringInStr($SafeHash, $Hash) Then
			_Error("Error starting restic.exe", Default, "Hash error - " & $Hash)
			Exit
		EndIf
	EndIf

	; Execute the restic command
	Local $Run = $ResticFullPath & " " & $Command

	_Log("  Command: " & $Run, 3)
	_UpdateEnv($aConfig)
	_Log("  Working Repository: " &	EnvGet("RESTIC_REPOSITORY"), 1)

	If $Command <> "stage env" Then
		_Log("  _RunWait - $RunSTDIO=" & $RunSTDIO, 2)
		Local $PID = _RunWait($Run, @ScriptDir, @SW_HIDE, $RunSTDIO, True, False)
		_UpdateEnv($aConfig, True) ; Remove env values
		_Log("")

		Return $PID
	EndIf

	Return ""
EndFunc   ;==>_Restic

; Do cleanup on script exit
Func _Exit()
	_Log("_Exit", 3)

	Global $SettingsForm
	GUIDelete($SettingsForm)

	; Close any instance of restic-browser
	If IsDeclared("ResticBrowserPid") Then ProcessClose($ResticBrowserPid)
	;ProcessClose("Restic-Browser.exe")

	; Delete any temp folders we ever created
	Local $sPath = @TempDir & "\"
	Local $aList = _FileListToArray($sPath, "sbr*.tmp", 2)
	If Not @error Then
		For $i = 1 To $aList[0]
			$RemovePath = $sPath & $aList[$i]
			DirRemove($RemovePath, 1)
			_Log("DirRemove: @error=" & @error & " (" & $aList[$i] & ")", 3)
		Next
	EndIf

	_Log("  Cleanup Done, Exiting Program")
EndFunc   ;==>_Exit

Func _IsGUIControlFocused($h_Wnd, $i_ControlID) ; Check if a control has focus.
    Return ControlGetHandle($h_Wnd, '', $i_ControlID) = ControlGetHandle($h_Wnd, '', ControlGetFocus($h_Wnd))
EndFunc   ;==>_OIG_IsFocused
