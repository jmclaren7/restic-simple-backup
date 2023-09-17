#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Change2CUI=y
#AutoIt3Wrapper_Res_Description=SimpleBackup
#AutoIt3Wrapper_Res_Fileversion=1.0.0.96
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_ProductVersion=1
#AutoIt3Wrapper_Res_LegalCopyright=SimpleBackup
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_requestedExecutionLevel=requireAdministrator
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

; Testing only, uncomment this to run as admin when running uncompiled
;#RequireAdmin

#include <Array.au3>
#Include <String.au3>
#include <StaticConstants.au3>
#include <AutoItConstants.au3>
#include <Crypt.au3>
#include <WinAPIDiag.au3>
#include <CommonFunctions.au3>
#include <GUIConstantsEx.au3>
#include <GuiEdit.au3>
#include <WindowsConstants.au3>

; Setup Logging For _ConsoleWrite
Global $LogToFile = 1
Global $LogFileMaxSize = 512
Global $LogLevel = 3

Global $Title = StringTrimRight(@ScriptName, 4)
_ConsoleWrite("Starting " & $Title)

Global $ResticFullPath = @TempDir & "\restic64.exe"
Global $ResticHash = "0x" & "dab3472f534e127b05b5c21e8edf2b8e0b79ae1c"
Global $HwKey = _WinAPI_UniqueHardwareID($UHID_MB) & DriveGetSerial(@HomeDrive & "\") & @CPUArch
Global $ConfigFile = StringTrimRight(@ScriptName, 4) & ".dat"
Global $ConfigFileFullPath = @ScriptDir & "\" & $ConfigFile
Global $Value_Prefix = "Config_"
Global $ValidEnvs =  "RESTIC_REPOSITORY|RESTIC_PASSWORD|AZURE_ACCOUNT_NAME|AZURE_ACCOUNT_KEY|AWS_ACCESS_KEY_ID|AWS_SECRET_ACCESS_KEY"
Global $ValidValues = "Setup_Password|Backup_Path|Backup_Prune" & "|" & $ValidEnvs


; Pack and unpack the restic executable
If FileInstall("restic64.exe", $ResticFullPath, 1) = 0 Then
	_ConsoleWrite("FileInstall error")
	Exit
Endif

; Register our exit function
OnAutoItExitRegister("_Exit")

; Load data from config and load to variables
_ReadConfig()

Global $SetupPassword = Eval($Value_Prefix & "Setup_Password")

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

		#Region ### START Koda GUI section ###
		$SettingsForm = GUICreate("Title", 679, 440, -1, -1, BitOR($GUI_SS_DEFAULT_GUI,$WS_SIZEBOX,$WS_THICKFRAME))
		$ApplyButton = GUICtrlCreateButton("Apply", 590, 408, 75, 25)
		GUICtrlSetResizing(-1, $GUI_DOCKRIGHT+$GUI_DOCKBOTTOM+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
		$ScriptEdit = GUICtrlCreateEdit("", 7, 3, 489, 401, BitOR($GUI_SS_DEFAULT_EDIT,$WS_BORDER), 0)
		GUICtrlSetData(-1, "")
		GUICtrlSetFont(-1, 10, 400, 0, "Consolas")
		GUICtrlSetResizing(-1, $GUI_DOCKLEFT+$GUI_DOCKRIGHT+$GUI_DOCKTOP+$GUI_DOCKBOTTOM)
		$CancelButton = GUICtrlCreateButton("Cancel", 505, 408, 75, 25)
		GUICtrlSetResizing(-1, $GUI_DOCKRIGHT+$GUI_DOCKBOTTOM+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
		$OKButton = GUICtrlCreateButton("OK", 419, 408, 75, 25)
		GUICtrlSetResizing(-1, $GUI_DOCKRIGHT+$GUI_DOCKBOTTOM+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
		$InitButton = GUICtrlCreateButton("Restic - Initialize Repository", 504, 8, 163, 33)
		$BackupButton = GUICtrlCreateButton("Restic - Backup", 504, 48, 163, 33)
		$ForgetButton = GUICtrlCreateButton("Restic - Forget/Prune", 504, 90, 163, 33)
		$CreateTaskButton = GUICtrlCreateButton("Create Scheduled Task", 504, 163, 163, 33)
		GUISetState(@SW_SHOW)
		#EndRegion ### END Koda GUI section ###

		WinSetTitle($SettingsForm, "", $Title)
		GUICtrlSetData($ScriptEdit, _MemoryToConfigRaw())

		While 1
			$nMsg = GUIGetMsg()
			Switch $nMsg
				Case $ApplyButton, $OKButton
					$GuiData = GUICtrlRead($ScriptEdit)
					_ConfigRawToMemory($GuiData)
					_WriteConfig()

					If $nMsg = $OKButton Then Exit

				Case $GUI_EVENT_CLOSE, $CancelButton
					Exit

				Case $InitButton
					_Restic("init")

				Case $BackupButton
					_Restic("backup """ & Eval($Value_Prefix & "Backup_Path") & """")

				Case $ForgetButton
					_Restic("forget --prune " & Eval($Value_Prefix & "Backup_Prune"))

				Case $CreateTaskButton
					$Run = "SCHTASKS /CREATE /SC DAILY /TN " & $Title & " /TR ""'" & @ScriptFullPath & "' backup"" /ST 22:00 /RL Highest /NP /F /RU System"
					_ConsoleWrite($Run)
					_RunWait($Run, @ScriptDir, @SW_SHOW, $STDIO_INHERIT_PARENT, True)

			EndSwitch
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
		If $InputPass = $SetupPassword Then ExitLoop

		$InputPass = InputBox($Title, "Enter Password", "", "*", Default, 130)
		If @error Then Exit

	Wend

	Return
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
Func _ConfigRawToMemory($RawData)
	_ConsoleWrite("_ConfigRawToMemory", 3)
	$RawData = StringSplit($RawData, @CRLF)

	For $o=1 To $RawData[0]
		$Key = StringLeft($RawData[$o], StringInStr($RawData[$o], "=") - 1)
		$KeyValue = StringTrimLeft($RawData[$o], StringInStr($RawData[$o], "="))

		Assign($Value_Prefix & $Key, $KeyValue, $ASSIGN_FORCEGLOBAL)
	Next

EndFunc

; Convert internal variables back into a string of key=value pairs
Func _MemoryToConfigRaw()
	_ConsoleWrite("_MemoryToConfigRaw", 3)

	Local $ConfigRaw

	$aValues = StringSplit($ValidValues, "|")

	For $o=1 To $aValues[0]
		$Key = $aValues[$o]
		$KeyData = Eval($Value_Prefix & $aValues[$o])

		$ConfigRaw &= $Key & "=" & $KeyData & @CRLF
	Next

	Return $ConfigRaw
EndFunc

; Read and decrypt the config file then load it to internal variables
Func _ReadConfig()
	_ConsoleWrite("_ReadConfig", 3)

	$FileData = FileRead($ConfigFileFullPath)

	; Decypt Data Here
	$FileData = BinaryToString(_Crypt_DecryptData($FileData, $HwKey, $CALG_AES_256))

	_ConfigRawToMemory($FileData)

	Return $FileData
EndFunc

; Convert internal variables to string, encrypt and write the config file
Func _WriteConfig()
	_ConsoleWrite("_WriteConfig", 3)

	$ConfigData = _MemoryToConfigRaw()

	; Encrypt
	$ConfigData = _Crypt_EncryptData($ConfigData, $HwKey, $CALG_AES_256)

	$hConfigFile = FileOpen($ConfigFileFullPath, 2)
	FileWrite($hConfigFile, $ConfigData)

	Return
EndFunc

; Execute a restic command
Func _Restic($Command)
	_ConsoleWrite("_Restic", 3)
	Local $Hash = _Crypt_HashFile($ResticFullPath, $CALG_SHA1)
	If $Hash <> $ResticHash Then
		_ConsoleWrite("Hash error - " & $Hash)
		Exit
	EndIf

	$Run = $ResticFullPath & " " & $Command

	;_ConsoleWrite("Command: " & $Run)
	_UpdateEnv()
	_ConsoleWrite("_RunWait", 2)
	_ConsoleWrite($Run, 3)
	_RunWait($Run, @ScriptDir, @SW_SHOW, $STDIO_INHERIT_PARENT, True)
	_ClearEnv()
EndFunc

Func _Exit()
	_ConsoleWrite("_Exit", 3)
	FileDelete($ResticFullPath)

EndFunc

