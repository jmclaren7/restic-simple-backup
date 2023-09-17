#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=SimpleBackup.ico
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Change2CUI=y
#AutoIt3Wrapper_Res_Description=SimpleBackup
#AutoIt3Wrapper_Res_Fileversion=1.0.0.171
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
#include <GUIConstantsEx.au3>
#include <GuiComboBox.au3>
#include <GuiEdit.au3>
#include <WindowsConstants.au3>

; https://github.com/jmclaren7/AutoITScripts/blob/master/CommonFunctions.au3
#include <CommonFunctions.au3>

; Setup Logging For _ConsoleWrite
Global $LogToFile = 1
Global $LogFileMaxSize = 512
Global $LogLevel = 3

; Setup some globals for general use
Global $Title = StringTrimRight(@ScriptName, 4)
_ConsoleWrite("Starting " & $Title)
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
If FileInstall("restic64.exe", $ResticFullPath, 1) = 0 Then
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
		_Auth()

		Global $SettingsForm, $RunCombo

		#Region ### START Koda GUI section ###
		$SettingsForm = GUICreate("Title", 507, 410, -1, -1, BitOR($GUI_SS_DEFAULT_GUI,$WS_SIZEBOX,$WS_THICKFRAME))
		$ApplyButton = GUICtrlCreateButton("Apply", 422, 376, 75, 25)
		GUICtrlSetResizing(-1, $GUI_DOCKRIGHT+$GUI_DOCKBOTTOM+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
		$ScriptEdit = GUICtrlCreateEdit("", 7, 3, 489, 321, BitOR($GUI_SS_DEFAULT_EDIT,$WS_BORDER), 0)
		GUICtrlSetData(-1, "")
		GUICtrlSetFont(-1, 10, 400, 0, "Consolas")
		GUICtrlSetResizing(-1, $GUI_DOCKLEFT+$GUI_DOCKRIGHT+$GUI_DOCKTOP+$GUI_DOCKBOTTOM)
		$CancelButton = GUICtrlCreateButton("Cancel", 337, 376, 75, 25)
		GUICtrlSetResizing(-1, $GUI_DOCKRIGHT+$GUI_DOCKBOTTOM+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
		$OKButton = GUICtrlCreateButton("OK", 251, 376, 75, 25)
		GUICtrlSetResizing(-1, $GUI_DOCKRIGHT+$GUI_DOCKBOTTOM+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
		$RunButton = GUICtrlCreateButton("Run", 444, 332, 51, 33)
		GUICtrlSetResizing(-1, $GUI_DOCKRIGHT+$GUI_DOCKBOTTOM+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
		$RunCombo = GUICtrlCreateCombo("Select or Type A Command", 15, 338, 417, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
		GUICtrlSetFont(-1, 9, 400, 0, "Consolas")
		GUICtrlSetResizing(-1, $GUI_DOCKLEFT+$GUI_DOCKRIGHT+$GUI_DOCKBOTTOM+$GUI_DOCKHEIGHT)
		$STDIOCheckBox = GUICtrlCreateCheckbox("Show More Output (Limits file log)", 16, 368, 225, 17)
		GUISetState(@SW_SHOW)
		#EndRegion ### END Koda GUI section ###

		; Set some of the GUI parameters so we don't have to do it directly in the form code
		WinSetTitle($SettingsForm, "", $Title)
		GUICtrlSetData($ScriptEdit, _MemoryToConfigRaw())
		_GUICtrlComboBox_SetDroppedWidth ( $RunCombo, 600)
		_UpdateCommandComboBox()
		GUICtrlSetState($STDIOCheckBox, $GUI_CHECKED)


		While 1
			; Continue based on GUI action
			$nMsg = GUIGetMsg()
			Switch $nMsg
				; Adjust the global used to determine STDIO streams for child processes
				Case $STDIOCheckBox
 					_ConsoleWrite("$STDIOCheckBox")
					If GUICtrlRead($STDIOCheckBox) = $GUI_CHECKED Then
						$RunSTDIO = $STDIO_INHERIT_PARENT
					Else
						$RunSTDIO = $STDERR_MERGED
					EndIf

				; Save or save and close
				Case $ApplyButton, $OKButton
					$GuiData = GUICtrlRead($ScriptEdit)
					_ConfigRawToMemory($GuiData)
					_WriteConfig()

					If $nMsg = $OKButton Then Exit

					_UpdateCommandComboBox()

				; Close program
				Case $GUI_EVENT_CLOSE, $CancelButton
					; Exit and run the registered exit function for cleanup
					Exit

				; Run the provided command
				Case $RunButton
					;
					GUISetState(@SW_DISABLE, $SettingsForm)
					WinSetTrans($SettingsForm, "", 180)

					; Don't continue if combo box is on the placeholder text
					If _GUICtrlComboBox_GetCurSel($RunCombo) = 0 Then ContinueLoop

					; Continue based on combobox value
					$RunComboText = GUICtrlRead($RunCombo)
					Switch $RunComboText
						Case "Create Scheduled Task"
							$Run = "SCHTASKS /CREATE /SC DAILY /TN " & $Title & " /TR ""'" & @ScriptFullPath & "' backup"" /ST 22:00 /RL Highest /NP /F /RU System"
							_ConsoleWrite($Run)
							_ConsoleWrite("")
							_RunWait($Run, @ScriptDir, @SW_SHOW, $RunSTDIO, True)

						Case "Restic Browser"
							; Pack and unpack the restic-browser executable
							DirCreate($TempDir)
							If FileInstall("Restic-Browser-Self.exe", $ResticBrowserFullPath, 1) = 0 Then
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

						Case Else
							_Restic($RunComboText)

					EndSwitch
			EndSwitch

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

; Prompt for a password before continuing
Func _Auth()
	_ConsoleWrite("_Auth", 3)

	Local $InputPass

	While 1
		; Check input first to deal with empty password
		If $InputPass = Eval($Value_Prefix & "Setup_Password") Then ExitLoop

		$InputPass = InputBox($Title, "Enter Password", "", "*", Default, 130)
		If @error Then Exit

	Wend

	Return
EndFunc

; Updates the options in the GUI combo box
Func _UpdateCommandComboBox()
	_ConsoleWrite("_UpdateCommandComboBox", 3)

	_GUICtrlComboBox_ResetContent ( $RunCombo )

	$Opts = "Select or type a command"
	$Opts &= "|" & "init  (Create the restic respository)"
	$Opts &= "|" & "Create Scheduled Task  (Automaticly creates a daily backup task for 10pm)"
	$Opts &= "|" & "Restic Browser  (Browse the repository to restore files)"
	$Opts &= "|" & "backup " & Eval($Value_Prefix & "Backup_Path") & "  (Runs a backup)"
	$Opts &= "|" & "forget --prune " & Eval($Value_Prefix & "Backup_Prune") & "  (Removes old backups)"
	$Opts &= "|" & "snapshots  (Lists snapshots in the repository)"
	$Opts &= "|" & "unlock  (Unlocks the repository in case restic had an issue)"
	$Opts &= "|" & "check --read-data  (Verifies all data in repo SLOW!!!)"
	$Opts &= "|" & "stats raw-data  (Show storage used)"
	$Opts &= "|" & "version  (Show restic version information)"
	$Opts &= "|" & "--help  (Show restic help information)"
	GUICtrlSetData($RunCombo, $Opts, "Select or type a command")

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
		$KeyData = Eval($Value_Prefix & $aValues[$o])

		$ConfigData &= $Key & "=" & $KeyData & @CRLF
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
		_ConsoleWrite("DirRemove: " & @error & " (" & $aList[$i] & ")")
	Next
EndFunc

