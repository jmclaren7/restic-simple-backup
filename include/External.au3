#include-once
#include <AutoItConstants.au3>
#Include <String.au3>

;===============================================================================
; Function Name:    _RunWait
; Description:		Improved version of RunWait that plays nice with my console logging
; Call With:		_RunWait($Run, $Working="")
; Parameter(s):
; Return Value(s):  On Success - Return value of Run() (Should be PID)
; 					On Failure - Return value of Run()
; Author(s):        JohnMC - JohnsCS.com
; Date/Version:		01/16/2016  --  v1.1
;===============================================================================
Func _RunWait($sProgram, $Working = "", $Show = @SW_HIDE, $Opt = $STDERR_MERGED, $Live = False)
	Local $sData, $iPid

	$iPid = Run($sProgram, $Working, $Show, $Opt)
	If @error Then
		_ConsoleWrite("_RunWait: Couldn't Run " & $sProgram)
		return SetError(1, 0, 0)
	endif

	$sData = _ProcessWaitClose($iPid, $Live)

	return SetError(0, $iPid, $sData)
endfunc
;===============================================================================
; Function Name:    _ProcessWaitClose
; Description:		ProcessWaitClose that handles stdout from the running process
;					Proccess must have been started with $STDERR_CHILD + $STDOUT_CHILD
; Call With:		_ProcessWaitClose($iPid)
; Parameter(s):
; Return Value(s):  On Success -
; 					On Failure -
; Author(s):        JohnMC - JohnsCS.com
; Date/Version:		09/8/2023  --  v1.3
;===============================================================================
Func _ProcessWaitClose($iPid, $Live = False, $Diag = False)
	Local $sData, $sStdRead

	While 1
		$sStdRead = StdoutRead($iPid)
		If @error Or $sStdRead = "" Then StderrRead($iPid)
		If @error And Not ProcessExists($iPid) Then ExitLoop
		$sStdRead = StringReplace($sStdRead, @CR&@LF&@CR&@LF, @CR&@LF)

		If $Diag Then
			$sStdRead = StringReplace($sStdRead, @CRLF, "@CRLF")
			$sStdRead = StringReplace($sStdRead, @CR, "@CR"&@CR)
			$sStdRead = StringReplace($sStdRead, @LF, "@LF"&@LF)
			$sStdRead = StringReplace($sStdRead, "@CRLF", "@CRLF"&@CRLF)
		EndIf

		If $sStdRead <> @CRLF Then
			$sData &= $sStdRead
			If $Live And $sStdRead <> "" Then
				If StringRight($sStdRead, 2) = @CRLF Then $sStdRead = StringTrimRight($sStdRead, 2)
				;If StringRight($sStdRead, 1) = @CR Then $sStdRead = StringTrimRight($sStdRead, 1) ; This may never be needed, leaving disabled
				If StringRight($sStdRead, 1) = @LF Then $sStdRead = StringTrimRight($sStdRead, 1)
				_ConsoleWrite($sStdRead)
			Endif
		Endif

		Sleep(5)
	WEnd

	return $sData
endfunc
;===============================================================================
; Function Name:   	_ConsoleWrite()
; Description:		Console & File Loging
; Call With:		_ConsoleWrite($Text,$SameLine)
; Parameter(s): 	$Text - Text to print
;					$Level - The level the given message *is*
;					$SameLine - (Optional) Will continue to print on the same line if set to 1, replace line if set to 2
;
; Return Value(s):  The Text Originaly Sent
; Notes:			Checks if global $LogToFile=1 or $CmdLineRaw contains "-debuglog" to see if log file should be writen
;					If Text = "OPENLOG" then log file is displayed (casesense)
; Author(s):        JohnMC - JohnsCS.com
; Date/Version:		05/05/2022 --  V1.3 Added $iSameLine option value "2" which will replace the current line
;					06/12/2019 --  v1.2 Added back -debuglog switch and updated notes
;					06/1/2012  --  v1.1
;===============================================================================
Func _ConsoleWrite($sMessage, $iLevel = 1, $iSameLine = 0)
	Local $hHandle, $sData

	if Eval("LogFilePath") = "" Then Global $LogFilePath = StringTrimRight(@ScriptFullPath,4)&"_Log.txt"
	if Eval("LogFileMaxSize") = "" Then Global $LogFileMaxSize = 0
	if Eval("LogToFile") = "" Then Global $LogToFile = False
	if StringInStr($CmdLineRaw, "-debuglog") Then Global $LogToFile = True
	if Eval("LogLevel") = "" Then Global $LogLevel = 3 ; The level of message to log - If no level set to 3

	If $sMessage == "OPENLOG" Then Return ShellExecute($LogFilePath)

	If $iLevel<=$LogLevel then
		$sMessage=StringReplace($sMessage,@CRLF&@CRLF,@CRLF) ;Remove Double CR
		If StringRight($sMessage,StringLen(@CRLF))=@CRLF Then $sMessage=StringTrimRight($sMessage,StringLen(@CRLF)) ; Remove last CR

		; Generate Timestamp
		Local $sTime=@YEAR&"-"&@MON&"-"&@MDAY&" "&@HOUR&":"&@MIN&":"&@SEC&"> "

		; Force CRLF
		$sMessage = StringRegExpReplace($sMessage, "((?<!\x0d)\x0a|\x0d(?!\x0a))", @CRLF)

		; Adds spaces for alignment after initial line
		$sMessage=StringReplace($sMessage,@CRLF,@CRLF&_StringRepeat(" ",StringLen($sTime)))

		If $iSameLine=0 then $sMessage=@CRLF&$sTime&$sMessage
		If $iSameLine=2 then $sMessage=@CR&$sTime&$sMessage

		ConsoleWrite($sMessage)

		If $LogToFile Then
			if $LogFileMaxSize<>0 AND FileGetSize($LogFilePath) > $LogFileMaxSize*1024 then
				$sMessage=FileRead($LogFilePath) & $sMessage
				$sMessage=StringTrimLeft($sMessage,StringInStr($sMessage, @CRLF, 0, 5))
				$hHandle=FileOpen($LogFilePath,2)
			Else
				$hHandle=FileOpen($LogFilePath,1)
			endif
			FileWrite($hHandle,$sMessage)
			FileClose($hHandle)

		endif
	endif

	Return $sMessage
EndFunc ;==> _ConsoleWrite
;===============================================================================
; Function Name:    _KeyValue()
; Description:		Work with 2d arrays treated as key value pairs such as the ones produced by INIReadSection()
; Call With:		_KeyValue(ByRef $Array, $Key[, $Value[, $Extended]])
; Parameter(s): 	$Array - A previously declared array, if not array, it will be made as one
;					$Key - The value to look for in the first column/dimention or the "Key" in an INI section
;		(Optional)	$Value - The value to write to the array
;		(Optional)	$Delete - If True, delete the specified key
;
; Return Value(s):  On Success - The value found or set or true if a value was deleted
; 					On Failure - "" and sets @error to 1
;
; Author(s):        JohnMC - JohnsCS.com
; Date/Version:		01/29/2010  --  v1.0
; Notes:            $Array[0][0] Contains the number of stored parameters
; Example:			_KeyValue($Settings, "trayicon", "1")
;===============================================================================
Func _KeyValue(ByRef $aArray, $Key, $Value = Default, $Delete = Default)
	Local $i

	If $Delete = Default Then $Delete = False

	; Make $Array an array if not already
	If Not IsArray($aArray) Then Dim $aArray[1][2]

	; Loop through array to check for existing key
	For $i = 1 To UBound($aArray) - 1
		If $aArray[$i][0] = $Key Then
			; Read existing value
			If $Value = Default Then
				Return $aArray[$i][1]

			; Update existing value
			Else
				$aArray[$i][1] = $Value
				$aArray[0][0] = UBound($aArray) - 1
				Return $Value
			EndIf

			; Delete existing value
			If $Delete Then
				Local $aNewArray[]
				; Loop through array and copy all keys/values not matching the specified key
				For $i = 1 To UBound($aArray) - 1
					; Skip the key to be deleted
					If $aArray[$i][0] = $Key Then ContinueLoop

					; Resize array and add new key/value
					ReDim $aArray[UBound($aNewArray) + 1][2]
					$aArray[UBound($aNewArray)][0] = $aArray[$i][0]
					$aArray[UBound($aNewArray)][1] = $aArray[$i][1]
				Next

				$aNewArray[0][0] = UBound($aArray) - 1

				; Return array with key/value removed
				$aArray = $aNewArray
				Return True
			EndIf
		EndIf
	Next

	; Add new key/value if it's been specified
	If $Value <> Default Then
		ReDim $aArray[UBound($aArray) + 1][2]
		$aArray[UBound($aArray) - 1][0] = $Key
		$aArray[UBound($aArray) - 1][1] = $Value
		$aArray[0][0] = UBound($aArray) - 1

		Return $Value
	endif

	; Return error because a key doesn't exist and nothing else to do
	SetError(1)
	Return ""
EndFunc ;==>_KeyValue

Func _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject = "", $as_Body = "", $s_Username = "", $s_Password = "", $s_CcAddress = "", $s_BccAddress = "", $IPPort = 587, $ssl = 0, $tls = True)
    Local $objEmail = ObjCreate("CDO.Message")
    $objEmail.From = '"' & $s_FromName & '" <' & $s_FromAddress & '>'
    $objEmail.To = $s_ToAddress
    Local $i_Error = 0
    Local $i_Error_desciption = ""
    If $s_CcAddress <> "" Then $objEmail.Cc = $s_CcAddress
    If $s_BccAddress <> "" Then $objEmail.Bcc = $s_BccAddress
    $objEmail.Subject = $s_Subject
    If StringInStr($as_Body, "<") And StringInStr($as_Body, ">") Then
        $objEmail.HTMLBody = $as_Body
    Else
        $objEmail.Textbody = $as_Body & @CRLF
    EndIf
    $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = $s_SmtpServer
    If Number($IPPort) = 0 then $IPPort = 25
    $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = $IPPort
    ;Authenticated SMTP
    If $s_Username <> "" Then
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = $s_Username
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = $s_Password
    EndIf
    ; Set security params
    If $ssl Then $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    If $tls Then $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendtls") = True
    ;Update settings
    $objEmail.Configuration.Fields.Update
    $objEmail.Fields.Update
    ; Sent the Message
    $objEmail.Send
    If @error Then
        SetError(2)
        Return 0
    EndIf
    $objEmail=""
EndFunc   ;==>_INetSmtpMailCom
