#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=include\SimpleBackup.ico
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Change2CUI=y
#AutoIt3Wrapper_Res_Description=SimpleBackup
#AutoIt3Wrapper_Res_Fileversion=1.0.0.229
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_ProductVersion=1
#AutoIt3Wrapper_Res_LegalCopyright=SimpleBackup
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_requestedExecutionLevel=highestAvailable
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

; Testing only, uncomment this to run as admin when running uncompiled
;#RequireAdmin

#include <Array.au3>
#include <File.au3>
#Include <String.au3>
#include <StaticConstants.au3>
#include <AutoItConstants.au3>
#include <Crypt.au3>
#include <WinAPIDiag.au3>
#include <WinAPIConv.au3>
#include <Inet.au3>
#include <GUIConstantsEx.au3>
#include <GuiComboBox.au3>
#include <GuiEdit.au3>
#include <GuiMenu.au3>
#include <WindowsConstants.au3>

; https://github.com/jmclaren7/autoit-scripts/blob/master/CommonFunctions.au3
#include <include\External.au3>

; Setup Logging For _ConsoleWrite
Global $LogToFile = 1
Global $LogFileMaxSize = 512
Global $LogLevel = 1
If Not @Compiled Then $LogLevel = 3

; Setup version and program title globals
Global $Version = 0
If @Compiled Then $Version = FileGetVersion(@AutoItExe)
Global $Title = StringTrimRight(@ScriptName, 4)
Global $TitleVersion = $Title & " v" & StringTrimLeft($Version, StringInStr($Version,".", 0, -1))
_ConsoleWrite("Starting " & $TitleVersion)

; Setup some miscellaneous globals
Global $TempDir = _TempFile (@TempDir, "sbr", "tmp", 10)
Global $ResticFullPath = $TempDir & "\restic.exe"
Global $ResticBrowserFullPath = $TempDir & "\Restic-Browser.exe"
Global $SMTPSettings = "Backup_Name|SMTP_Server|SMTP_UserName|SMTP_Password|SMTP_FromAddress|SMTP_FromName|SMTP_ToAddress|SMTP_SendOnFailure|SMTP_SendOnSuccess"
Global $InternalSettings = "Setup_Password|Backup_Path|Backup_Prune|" & $SMTPSettings
Global $RequiredSettings = "Setup_Password|Backup_Path|Backup_Prune|RESTIC_REPOSITORY|RESTIC_PASSWORD"
Global $SettingsTemplate = $RequiredSettings & "|AWS_ACCESS_KEY_ID|AWS_SECRET_ACCESS_KEY|RESTIC_READ_CONCURRENCY=4|RESTIC_PACK_SIZE=32|" & $SMTPSettings
Global $ConfigFile = StringTrimRight(@ScriptName, 4) & ".dat"
Global $ConfigFileFullPath = @ScriptDir & "\" & $ConfigFile
Global $ActiveProfile = "Default"

; $RunSTDIO will determine how we execute restic, if this is not done properly we will miss console output or log output
; This value changes depending on program contexts to give us the most aplicable output
Global $RunSTDIO = $STDERR_MERGED

; These hashes are used to verify the binaries right before they run
Global $ResticHash = "0x" & "dab3472f534e127b05b5c21e8edf2b8e0b79ae1c"
Global $ResticBrowserHash = "0x" & "6b6634710ff5011ace07666de838ad5c272e3d65"

; This key is used to encrypt the configuration file but is mostly just going to limit non-targeted/low-effort attacks, customizing they key for your own deployment could help though
Global $HwKey = _WinAPI_UniqueHardwareID($UHID_MB) & DriveGetSerial(@HomeDrive & "\") & @CPUArch

; Register our exit function for cleanup
OnAutoItExitRegister("_Exit")

; Pack and unpack the Restic executable
DirCreate($TempDir)
If FileInstall("include\restic64.exe", $ResticFullPath, 1) = 0 Then
	_ConsoleWrite("FileInstall error")
	Exit
Endif

; Interpret command line parameters
$Command = Default
For $i = 1 To $CmdLine[0]
	_ConsoleWrite("Parameter: " & $CmdLine[$i])
	Switch $CmdLine[$i]
		Case "profile"
			; The profile parameter and the following parameter are used to adjust the
			$i = $i + 1
			$ConfigFileFullPath = StringTrimRight($ConfigFileFullPath, 4) & "." & $CmdLine[$i] & ".dat"
			$ActiveProfile = $CmdLine[$i]
			_ConsoleWrite("Config file is now " & StringTrimLeft($ConfigFileFullPath, StringInStr($ConfigFileFullPath, "\", 0, -1)))

		Case Else
			; The first parameter not specialy handled must be the command
			If $Command = Default Then $Command = $CmdLine[$i]

	EndSwitch
Next

; Default command for when no parameters are provided
If $Command = Default Then $Command = "setup"
_ConsoleWrite("Command: " & $Command)

; This loop is used to effectively restart the program for when we switch profiles
While 1
	; Load config from file and load to array
	Global $aConfig = _ConfigToArray(_ReadConfig())

	; If config is empty load it with the template, otherwise make sure it at least has required settings
	if UBound($aConfig) = 0 Then
		_ForceRequiredConfig($aConfig, $SettingsTemplate)
	Else
		_ForceRequiredConfig($aConfig, $RequiredSettings)
	EndIf

	; Continue based on command line or default command
	Switch $Command
		; Display help information
		Case "help", "/?"
			_ConsoleWrite("Valid Restic Commands: version, stats, init, check, snapshots, backup, --help")
			_ConsoleWrite("Valid Script Commands: setup, command")

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

				_ConsoleWrite("Sending Email " & $sSubject)
				_INetSmtpMailCom(_KeyValue($aConfig, "SMTP_Server"), _KeyValue($aConfig, "SMTP_FromName"), _KeyValue($aConfig, "SMTP_FromAddress"), _
					_KeyValue($aConfig, "SMTP_ToAddress"), $sSubject, $sBody, _KeyValue($aConfig, "SMTP_UserName"), _KeyValue($aConfig, "SMTP_Password"))

			EndIf

			_Restic("forget --prune " & _KeyValue($aConfig, "Backup_Prune"))

		; Setup GUI
		Case "setup"
			WinMove("[TITLE:" & @AutoItExe & "; CLASS:ConsoleWindowClass]", "", 4, 4)

			_Auth()

			Global $SettingsForm, $RunCombo

			#Region ### START Koda GUI section ###
			$SettingsForm = GUICreate("Title", 601, 533, -1, -1, BitOR($GUI_SS_DEFAULT_GUI,$WS_SIZEBOX,$WS_THICKFRAME))
			$ApplyButton = GUICtrlCreateButton("Apply", 510, 496, 75, 25)
			GUICtrlSetResizing(-1, $GUI_DOCKRIGHT+$GUI_DOCKBOTTOM+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
			$ScriptEdit = GUICtrlCreateEdit("", 7, 3, 585, 441, BitOR($GUI_SS_DEFAULT_EDIT,$WS_BORDER), 0)
			GUICtrlSetData(-1, "")
			GUICtrlSetFont(-1, 10, 400, 0, "Consolas")
			GUICtrlSetResizing(-1, $GUI_DOCKLEFT+$GUI_DOCKRIGHT+$GUI_DOCKTOP+$GUI_DOCKBOTTOM)
			$CancelButton = GUICtrlCreateButton("Cancel", 425, 496, 75, 25)
			GUICtrlSetResizing(-1, $GUI_DOCKRIGHT+$GUI_DOCKBOTTOM+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
			$OKButton = GUICtrlCreateButton("OK", 339, 496, 75, 25)
			GUICtrlSetResizing(-1, $GUI_DOCKRIGHT+$GUI_DOCKBOTTOM+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
			$RunButton = GUICtrlCreateButton("Run", 532, 456, 51, 33)
			GUICtrlSetResizing(-1, $GUI_DOCKRIGHT+$GUI_DOCKBOTTOM+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
			$RunCombo = GUICtrlCreateCombo("Select or Type A Command", 15, 462, 505, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
			GUICtrlSetFont(-1, 9, 400, 0, "Consolas")
			GUICtrlSetResizing(-1, $GUI_DOCKLEFT+$GUI_DOCKRIGHT+$GUI_DOCKBOTTOM+$GUI_DOCKHEIGHT)
			GUISetState(@SW_SHOW)
			#EndRegion ### END Koda GUI section ###

			; Set some of the GUI parameters that we don't or can't do in Koda
			;WinMove($SettingsForm, "", Default, Default, 600, 450) ; Resize the window
			WinSetTitle($SettingsForm, "", $TitleVersion) ; Set the title from title variable
			If $ActiveProfile <> "Default" Then WinSetTitle($SettingsForm, "", $TitleVersion & " - " & $ActiveProfile) ; Change title to include profile if not the default profile
			GUICtrlSetData($ScriptEdit, _ArrayToConfig($aConfig)); Load the edit box with config data
			_GUICtrlComboBox_SetDroppedWidth($RunCombo, 600) ; Set the width of the combobox drop down beyond the width of the combobox
			_UpdateCommandComboBox() ; Set the options in the combobox

			; Setup a custom menu
			Global $MenuMsg = 0
			If Not IsDeclared("ExitMenuItem") Then Global Enum $ExitMenuItem = 1000, $ScheduledTaskMenuItem, $FixConsoleMenuItem, $BrowserMenuItem, _
				$VerboseMenuItem, $TemplateMenuItem, $NewProfileMenuItem, $AboutMenuItem, $WebsiteMenuItem

			; Setup file menu
			$g_hFile = _GUICtrlMenu_CreateMenu()
			; Create menu items based on config files found
			Global $aConfigFiles = _FileListToArray(@ScriptDir, StringReplace($ConfigFile, ".dat", "*.dat"), 1, True)
			For $i = 1 To Ubound($aConfigFiles) - 1
				$ProfileName = _GetProfileFromFullPath($aConfigFiles[$i])
				If $ProfileName = "" Then $ProfileName = $ProfileName & "Default"
				_ConsoleWrite("Found profile: "&$ProfileName)

				$cmdID = 1100 + $i

				If $ActiveProfile = $ProfileName Then $ProfileName = $ProfileName & " (Current Profile)"
				_GUICtrlMenu_InsertMenuItem($g_hFile, -1, $ProfileName, $cmdID)
				If $aConfigFiles[$i] = $ConfigFileFullPath Then _GUICtrlMenu_SetItemDisabled($g_hFile, $cmdID, True, False)
			Next
			_GUICtrlMenu_InsertMenuItem($g_hFile, -1, "New Profile/Config...", $NewProfileMenuItem)
			_GUICtrlMenu_InsertMenuItem($g_hFile, -1, "")
			_GUICtrlMenu_InsertMenuItem($g_hFile, -1, "Exit", $ExitMenuItem)

			; Setup tools menu
			$g_hTools = _GUICtrlMenu_CreateMenu()
			_GUICtrlMenu_InsertMenuItem($g_hTools, -1, "Create/Reset Scheduled Task (Profile: " & $ActiveProfile & ")", $ScheduledTaskMenuItem)
			_GUICtrlMenu_InsertMenuItem($g_hTools, -1, "Open Restic Browser", $BrowserMenuItem)
			_GUICtrlMenu_InsertMenuItem($g_hTools, -1, "Add Missing Configuration Options From Template", $TemplateMenuItem)

			; Setup advanced menu
			$g_hAdvanced = _GUICtrlMenu_CreateMenu()
			_GUICtrlMenu_InsertMenuItem($g_hAdvanced, -1, "Fix Console Live Output While In GUI (Breaks file log)", $FixConsoleMenuItem)
			_GUICtrlMenu_InsertMenuItem($g_hAdvanced, -1, "Verbose Logs (While In GUI)", $VerboseMenuItem)

			; Setup help menu
			$g_hHelp = _GUICtrlMenu_CreateMenu()
			_GUICtrlMenu_InsertMenuItem($g_hHelp, -1, "Visit Website", $WebsiteMenuItem)
			_GUICtrlMenu_InsertMenuItem($g_hHelp, -1, "About " & $Title, $AboutMenuItem)

			; Setup main menu
			$g_hMain = _GUICtrlMenu_CreateMenu(BitOr($MNS_CHECKORBMP, $MNS_MODELESS))
			_GUICtrlMenu_InsertMenuItem($g_hMain, -1, "&File", 0, $g_hFile)
			_GUICtrlMenu_InsertMenuItem($g_hMain, -1, "&Tools", 0, $g_hTools)
			_GUICtrlMenu_InsertMenuItem($g_hMain, -1, "&Advanced", 0, $g_hAdvanced)
			_GUICtrlMenu_InsertMenuItem($g_hMain, -1, "&Help", 0, $g_hHelp)

			; Create menu
			_GUICtrlMenu_SetMenu($SettingsForm, $g_hMain)

			; Additonal menu setup
			_GUICtrlMenu_SetItemState($g_hMain, $FixConsoleMenuItem, $MFS_CHECKED, True, False) ;
			If $LogLevel = 3 Then _GUICtrlMenu_SetItemState($g_hMain, $VerboseMenuItem, $MFS_CHECKED, True, False)
			GUIRegisterMsg($WM_COMMAND, "_WM_COMMAND")

			; GUI loop
			While 1
				$nMsg = GUIGetMsg()
				; Add menu actions from custom menu gui
				If $nMsg = 0 And $MenuMsg <> 0 Then
					$nMsg = $MenuMsg
					$MenuMsg = 0
				Endif
				If $nMsg <> 0 And Not ($nMsg > -13 And $nMsg < -3) Then _ConsoleWrite("Merged $nMsg = " & $nMsg, 3)

				; Continue based on GUI action
				Switch $nMsg
					; Save or save and close
					Case $ApplyButton, $OKButton
						_ConsoleWrite("Apply/Ok")
						$aConfig = _ConfigToArray(GUICtrlRead($ScriptEdit))
						_WriteConfig(_ArrayToConfig($aConfig))

						; Warn user if the backup path doesn't make sense
						$ErrorMessage = "Back_Path might contain an invalid path, your settings have been saved but please verify you have specified a valid path and used quotes properly. This warning was triggered based on the follow text..."
						$sBackup_Path = _KeyValue($aConfig, "Backup_Path")
						$aBackup_Path = StringRegExp($sBackup_Path, '["''].*?["'']', $STR_REGEXPARRAYGLOBALMATCH)

						; Check for paths inside quoted strings
						If IsArray($aBackup_Path) Then
							For $i = 0 To UBound($aBackup_Path) - 1
								$StrippedPath = StringReplace($aBackup_Path[$i], """", "")
								If Not FileExists($StrippedPath) Then MsgBox($MB_ICONINFORMATION, $TitleVersion, $ErrorMessage & @CRLF&@CRLF & $aBackup_Path[$i])

							Next
						; No quoted strings were found so check the entire string
						Else
							$StrippedPath = StringReplace($sBackup_Path, """", "") ; todo: only trim quotes from first and last
							If Not FileExists($StrippedPath) Then
								; The entire string wasn't a path so test up to the first space
								; This wont capture issues if we have more than one path seperated by space but that would be hard to do if we also want to access parameters here
								$StrippedPath = StringLeft($StrippedPath, StringInStr($StrippedPath, " "))
								If Not FileExists($StrippedPath) Then MsgBox($MB_ICONINFORMATION, $TitleVersion, $ErrorMessage & @CRLF&@CRLF & $sBackup_Path)
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
						If  _GUICtrlMenu_GetItemChecked($g_hMain, $nMsg, False) Then
							_GUICtrlMenu_SetItemChecked($g_hMain, $nMsg , False, False)
						Else
							_GUICtrlMenu_SetItemChecked($g_hMain, $nMsg , True, False)
						EndIf

					; Add missing key=value pairs to existing config
					Case $TemplateMenuItem
						_ForceRequiredConfig($aConfig, $SettingsTemplate)
						GUICtrlSetData($ScriptEdit, _ArrayToConfig($aConfig)); Load the edit box with config data

					; Open the github page
					Case $WebsiteMenuItem
						ShellExecute("https://github.com/jmclaren7/restic-simple-backup")

					; Open a dialog with program information
					Case $AboutMenuItem
						MsgBox($MB_ICONINFORMATION, $TitleVersion, _
							"Restic SimpleBackup" & @CRLF & "https://github.com/jmclaren7/restic-simple-backup" & @CRLF & "Copyright (c) 2023, John McLaren" & @CRLF & @CRLF & _
							"Restic" & @CRLF & "https://github.com/restic/restic" & @CRLF & "Copyright (c) 2014, Alexander Neumann" & @CRLF & @CRLF & _
							"Restic Browser" & @CRLF & "https://github.com/emuell/restic-browser" & @CRLF & "Copyright (c) 2022 Eduard Müller / taktik")

					; Create or switch profile
					Case $NewProfileMenuItem, 1100 To 1199
						_ConsoleWrite("Profile create/switch")

						If $nMsg=$NewProfileMenuItem Then
							; Prompt for profile name and convert to full path
							; If the profile name already exists the existing profile will load and wont be overwriten unless the user does so
							$NewProfile = InputBox($TitleVersion, "Enter a name for the new profile/config", "", "", Default, 130)
							If @error Then ContinueLoop
							$ConfigFileFullPath = StringTrimRight($ConfigFileFullPath, 4) & "." & $NewProfile & ".dat"
						Else
							$ConfigFileFullPath = $aConfigFiles[$nMsg - 1100]

						Endif

						_ConsoleWrite("$ConfigFileFullPath=" & $ConfigFileFullPath, 3)

						$ActiveProfile = _GetProfileFromFullPath($ConfigFileFullPath)

						_ConsoleWrite("$ActiveProfile=" & $ActiveProfile, 3)

						; Delete the GUI and restart
						GUIDelete($SettingsForm)
						ContinueLoop 2

					; Start the Restic-Browser
					Case $BrowserMenuItem
						; Pack and unpack the Restic-Browser executable
						DirCreate($TempDir)
						If FileInstall("include\Restic-Browser-Self.exe", $ResticBrowserFullPath, 1) = 0 Then
							_ConsoleWrite("FileInstall error")
							Exit
						Endif

						; Update PATH env so that Restic-browser.exe can start restic.exe
						$EnvPath = EnvGet("Path")
						If Not StringInStr($EnvPath, $TempDir) Then
							EnvSet("Path", $TempDir & ";" & $EnvPath)
							_ConsoleWrite("EnvSet: "&@error, 3)
						EndIf

						; Verify the hash of the restic-browser.exe
						Local $Hash = _Crypt_HashFile($ResticBrowserFullPath, $CALG_SHA1)
						If $Hash <> $ResticBrowserHash Then
							_ConsoleWrite("Hash error - " & $Hash)
							Exit
						EndIf

						; Load the Restic credential envs and start restic-browser.exe
						_UpdateEnv($aConfig)
						$ResticBrowserPid = Run($ResticBrowserFullPath)

					; Create a sceduled task to run the backup
					Case $ScheduledTaskMenuItem
						If $ActiveProfile <> "Default" Then
							$ProfileSwitch = " profile " & $ActiveProfile
							$TaskName = "." & $ActiveProfile
						Else
							$ProfileSwitch = ""
							$TaskName = ""
						EndIf

						$Run = "SCHTASKS /CREATE /SC DAILY /TN " & $Title & $TaskName  & " /TR ""'" & @ScriptFullPath & "' backup" & $ProfileSwitch & """ /ST 22:00 /RL Highest /NP /F /RU System"
						_ConsoleWrite($Run)
						$Return = _RunWait($Run, @ScriptDir, @SW_SHOW, $STDERR_MERGED, True)
						If StringInStr($Return, "SUCCESS: ") Then
							MsgBox(0, $TitleVersion, "Scheduled task created. Please review and test the task.")
						Else
							MsgBox($MB_ICONERROR, $TitleVersion, "Error creating scheduled task.")
							If Not @Compiled Then MsgBox($MB_ICONERROR, $TitleVersion, "Are you running without admin?")
						EndIf

					; Run the command provided from the combobox
					Case $RunButton
						; Adjust window visibility and activation due to long running process
						GUISetState(@SW_DISABLE, $SettingsForm)
						WinSetTrans($SettingsForm, "", 180)
						If @Compiled Then WinActivate("[TITLE:" & @AutoItExe & "; CLASS:ConsoleWindowClass]")

						; Don't continue if combo box is on the placeholder text
						If _GUICtrlComboBox_GetCurSel($RunCombo) = 0 Then ContinueLoop

						; Continue based on combobox value
						$RunComboText = GUICtrlRead($RunCombo)
						Switch $RunComboText
							Case "custom/advanced commands can be added here in the future"

							Case "Test Email"
								If _KeyValue($aConfig, "SMTP_Server") Then
									$sSubject = "Test Subject"
									$sBody = "Test Body"
									$Return = _INetSmtpMailCom(_KeyValue($aConfig, "SMTP_Server"), _KeyValue($aConfig, "SMTP_FromName"), _KeyValue($aConfig, "SMTP_FromAddress"), _KeyValue($aConfig, "SMTP_ToAddress"), _
										$sSubject, $sBody, _KeyValue($aConfig, "SMTP_UserName"), _KeyValue($aConfig, "SMTP_Password"))
									_ConsoleWrite("Email Test: $Return=" & $Return & "  @error=" & @error)
								EndIf

							Case Else
								_Restic($RunComboText)

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
					$RunSTDIO = $STDIO_INHERIT_PARENT
				Else
					$RunSTDIO = $STDERR_MERGED
				EndIf

				; Re-enable the window when the script resumes from a blocking task
				If BitAND(WinGetState($SettingsForm), @SW_DISABLE) Then
					GUISetState(@SW_ENABLE, $SettingsForm)
					WinSetTrans($SettingsForm, "", 255)
				Endif

				; Removes help text from combobox selection
				$RunComboText = GUICtrlRead($RunCombo)
				If StringRight($RunComboText, 1) = ")" And StringInStr($RunComboText, "  (") And Not _GUICtrlComboBox_GetDroppedState($RunCombo) Then
					$NewText = StringLeft($RunComboText, StringInStr($RunComboText, "  (") -1)
					_GUICtrlComboBox_SetEditText($RunCombo, $NewText)
				Endif

				Sleep(10)
			WEnd

		Case Else
			_ConsoleWrite("Invalid Command")

 	EndSwitch

	Exit ; To support existing program flow since adding loop used to restart GUI
Wend

;=====================================================================================
;=====================================================================================
; Extract profile name from full path
Func _GetProfileFromFullPath($sPath)
	Local $Return = StringRegExp($sPath, "\" & $Title & ".([0-9a-zA-Z.-_ ]+).dat", 1)
	If @error Then
		Return "Default"
	Else
		Return $Return[0]
	EndIf
EndFunc

; Special function to handle messages from custom GUI menu
Func _WM_COMMAND($hWnd, $iMsg, $wParam, $lParam)
		Local $Temp = _WinAPI_LoWord($wParam)

		;_ConsoleWrite("_WM_COMMAND ($wParam = " & $Temp & ") ", 3)

		If $Temp >= 1000 And $Temp < 1999 Then
			_ConsoleWrite("Updated $MenuMsg To " & $Temp, 3)
			Global $MenuMsg = $Temp
		EndIf

        Return $GUI_RUNDEFMSG
EndFunc

; Prompt for a password before continuing
Func _Auth()
	_ConsoleWrite("_Auth", 3)

	Local $InputPass

	While 1
		; Check input first to deal with empty password
		If $InputPass = _KeyValue($aConfig, "Setup_Password") Then ExitLoop

		$InputPass = InputBox($TitleVersion, "Enter Password", "", "*", Default, 130)
		If @error Then Exit

	Wend

	Return
EndFunc

; Updates the options in the GUI combo box
Func _UpdateCommandComboBox()
	_ConsoleWrite("_UpdateCommandComboBox", 3)

	_GUICtrlComboBox_ResetContent ( $RunCombo )

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

EndFunc

; Update the environmental variables
Func _UpdateEnv($aArray, $Delete = Default)
	_ConsoleWrite("_UpdateEnv", 3)

	If $Delete = Default Then $Delete = False

	Local $aInternalSettings = StringSplit($InternalSettings, "|")

	; Loop array and set or delete Env
	For $i=1 To $aArray[0][0]
		; Skip If the value is a comment, empty or internal
		If StringLeft($aArray[$i][0], 1) = "#" Then ContinueLoop
		If $aArray[$i][0] = "" OR $aArray[$i][1] = "" Then ContinueLoop
		If _ArraySearch ($aInternalSettings, $aArray[$i][0], 0, 0, 0, 0) <> -1 Then ContinueLoop

		If $Delete Then
			_ConsoleWrite("  EnvSet (Delete): " & $aArray[$i][0], 3)
			EnvSet($aArray[$i][0], "")
		Else
			_ConsoleWrite("  EnvSet: " & $aArray[$i][0], 3)
			EnvSet($aArray[$i][0], $aArray[$i][1])
		EndIf
	Next

	Return
EndFunc

; Force required keys to show in the config array
Func _ForceRequiredConfig(Byref $aArray, $sRequired)
	Local $aKeyValuem, $sValue

	$sRequired = StringSplit($sRequired, "|")

	For $i = 1 To $sRequired[0]
		$aKeyValue = StringSplit($sRequired[$i],"=")
		_KeyValue($aArray, $aKeyValue[1])

		If $aKeyValue[0] = 2 Then
			$sValue = $aKeyValue[2]
		Else
			$sValue = ""
		EndIf

		If @error Then _KeyValue($aArray, $aKeyValue[1], $sValue)
	Next
EndFunc

; Converts string of key=value pairs to array and handles comments
Func _ConfigToArray($ConfigData)
	_ConsoleWrite("_ConfigToArray", 3)

	$aConfigLines = StringSplit($ConfigData, @CRLF, 1)
	Local $aArray[]

	; Loop through each line
	For $i = 1 To $aConfigLines[0]
		; Prefix comments with #<UID>= so that we can treat them as a key=value pair, each with a unique key
		If StringLeft($aConfigLines[$i], 1) = "#" Then $aConfigLines[$i] = "#" & $i & "=" & $aConfigLines[$i]

		$Key = StringLeft($aConfigLines[$i], StringInStr($aConfigLines[$i], "=") - 1)
		$Key = StringStripWS ($Key, 1 + 2)

		If $Key = "" Then ContinueLoop

		$KeyValue = StringTrimLeft($aConfigLines[$i], StringInStr($aConfigLines[$i], "="))
		$KeyValue = StringStripWS($KeyValue, 1 + 2)

		_KeyValue($aArray, $Key, $KeyValue)
	Next


	Return $aArray
EndFunc

; Converts array back to a string of key=value pairs and fix comments
Func _ArrayToConfig($aArray)
	_ConsoleWrite("_ArrayToConfig", 3)

	If UBound($aArray, 0) <> 2 Then Return SetError(1, 0, "")

	Local $ConfigData, $Add

	; Loop array and combine key=value pairs
	For $i=1 To $aArray[0][0]
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
EndFunc

; Read and decrypt the config file
Func _ReadConfig()
	_ConsoleWrite("_ReadConfig", 3)

	Local $ConfigData = FileRead($ConfigFileFullPath)

	; Decypt Data Here
	$ConfigData = BinaryToString(_Crypt_DecryptData($ConfigData, $HwKey, $CALG_AES_256))

	Return $ConfigData
EndFunc

; Encrypt and write the config file
Func _WriteConfig($ConfigData)
	_ConsoleWrite("_WriteConfig", 3)

	; Encrypt
	$ConfigData = _Crypt_EncryptData($ConfigData, $HwKey, $CALG_AES_256)

	$hConfigFile = FileOpen($ConfigFileFullPath, 2)
	FileWrite($hConfigFile, $ConfigData)

	Return
EndFunc

; Execute a Restic command
Func _Restic($Command, $Opt = $RunSTDIO)
	_ConsoleWrite("_Restic", 3)

	Local $Hash = _Crypt_HashFile($ResticFullPath, $CALG_SHA1)

	If $Hash <> $ResticHash Then
		_ConsoleWrite("Hash error - " & $Hash)
		Exit
	EndIf

	Local $Run = $ResticFullPath & " " & $Command

	_ConsoleWrite("  Command: " & $Run, 3)
	_UpdateEnv($aConfig)
	_ConsoleWrite("  Working Repository: " & EnvGet("RESTIC_REPOSITORY"), 1)

	_ConsoleWrite("  _RunWait", 2)
	If $Opt = $STDIO_INHERIT_PARENT Then _ConsoleWrite("")
	Local $Return = _RunWait($Run, @ScriptDir, @SW_Hide, $Opt, True)
	_UpdateEnv($aConfig, True) ; Remove env values

	Return $Return
EndFunc

Func _Exit()
	_ConsoleWrite("_Exit", 3)

	GUIDelete($SettingsForm)

	; Close any instance of restic-browser
	If IsDeclared("ResticBrowserPid") Then ProcessClose($ResticBrowserPid)
	ProcessClose("Restic-Browser.exe")

	; Delete any temp folders we ever created
	Local $sPath = @TempDir & "\"
	Local $aList = _FileListToArray($sPath, "sbr*.tmp", 2)
	For $i = 1 To $aList[0]
		$RemovePath = $sPath & $aList[$i]
		DirRemove($RemovePath, 1)
		_ConsoleWrite("DirRemove: " & @error & " (" & $aList[$i] & ")", 3)
	Next

	_ConsoleWrite("  Cleanup Done, Exiting Program")
EndFunc

