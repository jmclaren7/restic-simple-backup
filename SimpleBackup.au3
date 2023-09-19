#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=include\SimpleBackup.ico
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Change2CUI=y
#AutoIt3Wrapper_Res_Description=SimpleBackup
#AutoIt3Wrapper_Res_Fileversion=1.0.0.200
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
#include <GUIConstantsEx.au3>
#include <GuiComboBox.au3>
#include <GuiEdit.au3>
#include <GuiMenu.au3>
#include <WindowsConstants.au3>

; https://github.com/jmclaren7/AutoITScripts/blob/master/CommonFunctions.au3
#include <include\External.au3>

; Setup Logging For _ConsoleWrite
Global $LogToFile = 1
Global $LogFileMaxSize = 512
Global $LogLevel = 1

; Setup some globals for general use
Global $Version = 0
If @Compiled Then $Version = FileGetVersion(@AutoItExe)
Global $Title = StringTrimRight(@ScriptName, 4)
Global $TitleVersion = $Title & " v" & StringTrimLeft($Version, StringInStr($Version,".", 0, -1))
_ConsoleWrite("Starting " & $TitleVersion)
Global $TempDir = _TempFile (@TempDir, "sbr", "tmp", 10)
Global $ResticFullPath = $TempDir & "\restic.exe"
Global $ResticBrowserFullPath = $TempDir & "\Restic-Browser.exe"
Global $ResticHash = "0x" & "dab3472f534e127b05b5c21e8edf2b8e0b79ae1c"
Global $ResticBrowserHash = "0x" & "6b6634710ff5011ace07666de838ad5c272e3d65"
Global $HwKey = _WinAPI_UniqueHardwareID($UHID_MB) & DriveGetSerial(@HomeDrive & "\") & @CPUArch
Global $ConfigFile = StringTrimRight(@ScriptName, 4) & ".dat"
Global $ConfigFileFullPath = @ScriptDir & "\" & $ConfigFile
Global $Value_Prefix = "Config_"
Global $ValidEnvs =  "RESTIC_REPOSITORY|RESTIC_PASSWORD|AZURE_ACCOUNT_NAME|AZURE_ACCOUNT_KEY|AWS_ACCESS_KEY_ID|AWS_SECRET_ACCESS_KEY"
Global $ValidValues = "Setup_Password|Backup_Path|Backup_Prune" & "|" & $ValidEnvs
Global $RunSTDIO = $STDERR_MERGED

; Pack and unpack the restic executable
DirCreate($TempDir)
If FileInstall("include\restic64.exe", $ResticFullPath, 1) = 0 Then
	_ConsoleWrite("FileInstall error")
	Exit
Endif

; Register our exit function for cleanup
OnAutoItExitRegister("_Exit")

; Load data from config and load to variables
_ReadConfig()

; Interpret command line parameters
If $CmdLine[0] >= 1 Then
	$Command = $CmdLine[1]
Else
	$Command = "setup"
EndIf

_ConsoleWrite("Command: " & $Command)

Switch $Command
	; Display help information
	Case "help", "/?"
		_ConsoleWrite("Valid Restic Commands: version, stats, init, check, snapshots, backup, --help")
		_ConsoleWrite("Valid Script Commands: setup, command")

	; Basic commands allowed to be passed to the restic executable
	Case "version", "stats", "init", "check", "snapshots", "--help"
		_Restic($Command)

	; Pass arbitrary commands to the restic executable
	Case "c", "command"
		_Auth()

		$Run = StringTrimLeft($CmdLineRaw, StringLen($CmdLine[1]) + 1)
		_Restic($Run)

	; Backup command
	Case "backup"
		_Restic("backup """ & Eval($Value_Prefix & "Backup_Path") & """")
		_Restic("forget --prune " & Eval($Value_Prefix & "Backup_Prune"))

	; Setup GUI
	Case "setup"
		WinMove("[TITLE:" & @AutoItExe & "; CLASS:ConsoleWindowClass]", "", 4, 4)

		_Auth()

		Global $SettingsForm, $RunCombo

		#Region ### START Koda GUI section ###
		$SettingsForm = GUICreate("Title", 601, 451, -1, -1, BitOR($GUI_SS_DEFAULT_GUI,$WS_SIZEBOX,$WS_THICKFRAME))
		$ApplyButton = GUICtrlCreateButton("Apply", 510, 416, 75, 25)
		GUICtrlSetResizing(-1, $GUI_DOCKRIGHT+$GUI_DOCKBOTTOM+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
		$ScriptEdit = GUICtrlCreateEdit("", 7, 3, 585, 361, BitOR($GUI_SS_DEFAULT_EDIT,$WS_BORDER), 0)
		GUICtrlSetData(-1, "")
		GUICtrlSetFont(-1, 10, 400, 0, "Consolas")
		GUICtrlSetResizing(-1, $GUI_DOCKLEFT+$GUI_DOCKRIGHT+$GUI_DOCKTOP+$GUI_DOCKBOTTOM)
		$CancelButton = GUICtrlCreateButton("Cancel", 425, 416, 75, 25)
		GUICtrlSetResizing(-1, $GUI_DOCKRIGHT+$GUI_DOCKBOTTOM+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
		$OKButton = GUICtrlCreateButton("OK", 339, 416, 75, 25)
		GUICtrlSetResizing(-1, $GUI_DOCKRIGHT+$GUI_DOCKBOTTOM+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
		$RunButton = GUICtrlCreateButton("Run", 532, 376, 51, 33)
		GUICtrlSetResizing(-1, $GUI_DOCKRIGHT+$GUI_DOCKBOTTOM+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
		$RunCombo = GUICtrlCreateCombo("Select or Type A Command", 15, 382, 505, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
		GUICtrlSetFont(-1, 9, 400, 0, "Consolas")
		GUICtrlSetResizing(-1, $GUI_DOCKLEFT+$GUI_DOCKRIGHT+$GUI_DOCKBOTTOM+$GUI_DOCKHEIGHT)
		GUISetState(@SW_SHOW)
		#EndRegion ### END Koda GUI section ###

		; Set some of the GUI parameters that we don't or can't do in Koda
		;WinMove($SettingsForm, "", Default, Default, 600, 450) ; Resize the window
		WinSetTitle($SettingsForm, "", $TitleVersion) ; Set the title from title variable
		GUICtrlSetData($ScriptEdit, _MemoryToConfigRaw()) ; Load the edit box with config data
		_GUICtrlComboBox_SetDroppedWidth($RunCombo, 600) ; Set the width of the combobox drop down beyond the width of the combobox
		_UpdateCommandComboBox() ; Set the options in the combobox

		Global $MenuMsg = 0
		Global Enum $ExitMenuItem = 1000, $ScheduledTaskMenuItem, $FixConsoleMenuItem, $BrowserMenuItem, $VerboseMenuItem
		; Create menus
		$g_hFile = _GUICtrlMenu_CreateMenu()
		_GUICtrlMenu_InsertMenuItem($g_hFile, 0, "Exit", $ExitMenuItem)
		$g_hTools = _GUICtrlMenu_CreateMenu()
		_GUICtrlMenu_InsertMenuItem($g_hTools, 0, "Create/Reset Scheduled Task", $ScheduledTaskMenuItem)
		_GUICtrlMenu_InsertMenuItem($g_hTools, 1, "Open Restic Browser", $BrowserMenuItem)
		$g_hAdvanced = _GUICtrlMenu_CreateMenu()
		_GUICtrlMenu_InsertMenuItem($g_hAdvanced, 0, "Fix Console Live Output While In GUI (Breaks file log)", $FixConsoleMenuItem)
		_GUICtrlMenu_InsertMenuItem($g_hAdvanced, 1, "Verbose Logs (While In GUI)", $VerboseMenuItem)
		; Create Main menu
		$g_hMain = _GUICtrlMenu_CreateMenu(BitOr($MNS_CHECKORBMP, $MNS_MODELESS)) ; ..for MNS_MODELESS, only this "main menu" is needed.
		_GUICtrlMenu_InsertMenuItem($g_hMain, 0, "&File", 0, $g_hFile)
		_GUICtrlMenu_InsertMenuItem($g_hMain, 1, "&Tools", 0, $g_hTools)
		_GUICtrlMenu_InsertMenuItem($g_hMain, 2, "&Advanced", 0, $g_hAdvanced)
		_GUICtrlMenu_SetMenu($SettingsForm, $g_hMain)

		_GUICtrlMenu_SetItemState($g_hMain, $FixConsoleMenuItem, $MFS_CHECKED, True, False)
		If $LogLevel = 3 Then _GUICtrlMenu_SetItemState($g_hMain, $VerboseMenuItem, $MFS_CHECKED, True, False)
		GUIRegisterMsg($WM_COMMAND, "_WM_COMMAND")


		While 1
			$nMsg = GUIGetMsg()
			; Add menu actions from custom menu gui
			If $nMsg = 0 And $MenuMsg <> 0 Then
				$nMsg = $MenuMsg
				$MenuMsg = 0
			Endif
			If $nMsg <> 0 And $nMsg <> -11 Then _ConsoleWrite("Merged $nMsg = "&$nMsg, 3)

			; Continue based on GUI action
			Switch $nMsg
				; Save or save and close
				Case $ApplyButton, $OKButton
					$GuiData = GUICtrlRead($ScriptEdit)
					_ConfigRawToMemory($GuiData)
					_WriteConfig()

					If $nMsg = $OKButton Then Exit

					_UpdateCommandComboBox()

				; Close program
				Case $GUI_EVENT_CLOSE, $CancelButton, $ExitMenuItem
					; Exit and run the registered exit function for cleanup
					Exit

				; Handle menu items that use checkboxes
				Case $FixConsoleMenuItem, $VerboseMenuItem
					If  _GUICtrlMenu_GetItemChecked($g_hMain, $nMsg, False) Then
						_GUICtrlMenu_SetItemChecked($g_hMain, $nMsg , False, False)
					Else
						_GUICtrlMenu_SetItemChecked($g_hMain, $nMsg , True, False)
					EndIf

				Case $BrowserMenuItem
					; Pack and unpack the restic-browser executable
					DirCreate($TempDir)
					If FileInstall("include\Restic-Browser-Self.exe", $ResticBrowserFullPath, 1) = 0 Then
						_ConsoleWrite("FileInstall error")
						Exit
					Endif

					; Update PATH env so that restic-browser.exe can start restic.exe
					$EnvPath = EnvGet("Path")
					If Not StringInStr($EnvPath, $TempDir) Then
						EnvSet("Path", $TempDir & ";" & $EnvPath)
						_ConsoleWrite("EnvSet: "&@error)
					EndIf

					; Verify the hash of the restic-browser.exe
					Local $Hash = _Crypt_HashFile($ResticBrowserFullPath, $CALG_SHA1)
					If $Hash <> $ResticBrowserHash Then
						_ConsoleWrite("Hash error - " & $Hash)
						Exit
					EndIf

					; Load the restic credential envs and start restic-browser.exe
					_UpdateEnv()
					$ResticBrowserPid = Run($ResticBrowserFullPath)

				Case $ScheduledTaskMenuItem
					$Run = "SCHTASKS /CREATE /SC DAILY /TN " & $Title & " /TR ""'" & @ScriptFullPath & "' backup"" /ST 22:00 /RL Highest /NP /F /RU System"
					_ConsoleWrite($Run)
					$Return = _RunWait($Run, @ScriptDir, @SW_SHOW, $STDERR_MERGED, True)
					If StringInStr($Return, "SUCCESS: ") Then
						MsgBox(0, $TitleVersion, "Scheduled task created. Please review and test the task.")
					Else
						MsgBox($MB_ICONERROR, $TitleVersion, "Error creating scheduled task.")
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
						Case "place holder"

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

;=====================================================================================
;=====================================================================================
; Special function to handle messages from custom gui menu
Func _WM_COMMAND($hWnd, $iMsg, $wParam, $lParam)
		Local $Temp = _WinAPI_LoWord($wParam)

		;_ConsoleWrite("_WM_COMMAND ($wParam = " & $Temp & ") ", 3)

		If $Temp >= 1000 And $Temp < 1100 Then
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
		If $InputPass = Eval($Value_Prefix & "Setup_Password") Then ExitLoop

		$InputPass = InputBox($TitleVersion, "Enter Password", "", "*", Default, 130)
		If @error Then Exit

	Wend

	Return
EndFunc

; Updates the options in the GUI combo box
Func _UpdateCommandComboBox()
	_ConsoleWrite("_UpdateCommandComboBox", 3)

	_GUICtrlComboBox_ResetContent ( $RunCombo )

	$Opts = "Select or type a restic command"
	$Opts &= "|" & "init  (Create the restic respository)"
	$Opts &= "|" & "backup " & Eval($Value_Prefix & "Backup_Path") & "  (Runs a backup)"
	$Opts &= "|" & "forget --prune " & Eval($Value_Prefix & "Backup_Prune") & "  (Removes old backups)"
	$Opts &= "|" & "snapshots  (Lists snapshots in the repository)"
	$Opts &= "|" & "unlock  (Unlocks the repository in case restic had an issue)"
	$Opts &= "|" & "check --read-data  (Verifies all data in repo SLOW!!!)"
	$Opts &= "|" & "stats raw-data  (Show storage used)"
	$Opts &= "|" & "version  (Show restic version information)"
	$Opts &= "|" & "--help  (Show restic help information)"
	GUICtrlSetData($RunCombo, $Opts, "Select or type a restic command")

EndFunc

; Update the enviromental variables from the valid list
Func _UpdateEnv()
	_ConsoleWrite("_UpdateEnv", 3)

	Local $aValues = StringSplit($ValidEnvs, "|")
	For $o=1 To $aValues[0]
		Local $Value = Eval($Value_Prefix & $aValues[$o])
		If $Value <> "" Then EnvSet($aValues[$o], $Value)

	Next

	Return
EndFunc

; Remove enviromental variables to try and prevent them from being read externally
Func _ClearEnv()
	_ConsoleWrite("_ClearEnv", 3)

	$aValues = StringSplit($ValidEnvs, "|")
	For $o=1 To $aValues[0]
		EnvSet($aValues[$o], "")

	Next

	Return

EndFunc

; Convert a tring of key=value pairs into internal variables
Func _ConfigRawToMemory($ConfigData)
	_ConsoleWrite("_ConfigRawToMemory", 3)

	$ConfigData = StringSplit($ConfigData, @CRLF)

	For $o=1 To $ConfigData[0]
		$Key = StringLeft($ConfigData[$o], StringInStr($ConfigData[$o], "=") - 1)
		$KeyValue = StringTrimLeft($ConfigData[$o], StringInStr($ConfigData[$o], "="))

		Assign($Value_Prefix & $Key, $KeyValue, $ASSIGN_FORCEGLOBAL)
	Next

EndFunc

; Convert internal variables back into a string of key=value pairs
Func _MemoryToConfigRaw()
	_ConsoleWrite("_MemoryToConfigRaw", 3)

	Local $ConfigData

	$aValues = StringSplit($ValidValues, "|")

	For $o=1 To $aValues[0]
		$Key = $aValues[$o]
		$KeyValue = Eval($Value_Prefix & $aValues[$o])

		$ConfigData &= $Key & "=" & $KeyValue & @CRLF
	Next

	Return $ConfigData
EndFunc

; Read and decrypt the config file then load it to internal variables
Func _ReadConfig()
	_ConsoleWrite("_ReadConfig", 3)

	Local $ConfigData = FileRead($ConfigFileFullPath)

	; Decypt Data Here
	$ConfigData = BinaryToString(_Crypt_DecryptData($ConfigData, $HwKey, $CALG_AES_256))

	_ConfigRawToMemory($ConfigData)

	Return $ConfigData
EndFunc

; Convert internal variables to string, encrypt and write the config file
Func _WriteConfig()
	_ConsoleWrite("_WriteConfig", 3)

	Local $ConfigData = _MemoryToConfigRaw()

	; Encrypt
	$ConfigData = _Crypt_EncryptData($ConfigData, $HwKey, $CALG_AES_256)

	$hConfigFile = FileOpen($ConfigFileFullPath, 2)
	FileWrite($hConfigFile, $ConfigData)

	Return
EndFunc

; Execute a restic command
Func _Restic($Command, $Opt = $RunSTDIO)
	_ConsoleWrite("_Restic", 3)

	Local $Hash = _Crypt_HashFile($ResticFullPath, $CALG_SHA1)

	If $Hash <> $ResticHash Then
		_ConsoleWrite("Hash error - " & $Hash)
		Exit
	EndIf

	Local $Run = $ResticFullPath & " " & $Command

	_ConsoleWrite("Command: " & $Run, 3)
	_UpdateEnv()
	_ConsoleWrite("_RunWait", 2)

	If $Opt = $STDIO_INHERIT_PARENT Then _ConsoleWrite("")
	_RunWait($Run, @ScriptDir, @SW_Hide, $Opt, True)
	_ClearEnv()

EndFunc

Func _Exit()
	_ConsoleWrite("_Exit", 3)

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

	_ConsoleWrite("Cleanup Done, Exiting Program")
EndFunc

