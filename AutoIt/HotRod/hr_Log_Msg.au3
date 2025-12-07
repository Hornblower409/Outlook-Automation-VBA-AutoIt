#include-once

#include "HotRod\hr_Error_Exit.au3"

; =====================================================================
; 2023-11-30 - hr_Log_Msg - Write a line to the HotRod log file
;
;	Usage:
;
;	hr_Log_Msg( $ThisFunc, "StepName", @error, "MsgText", "$P1Name" )
;	hr_Log_Msg( $ThisFunc, "StepName", @error, "MsgText", "$P1Name", $P1Value, ... )
;
; =====================================================================

; TEST - hr_Log_Msg( "Test", "StepName", "-409", "MsgText", "$P1Name", "$P1Value" )

Func hr_Log_Msg( _
	Const $FuncName = Default, _
	$StepName = Default, _
	$ErrorValue = Default, _
	$MsgText = Default, _
	$P1Name = Default, $P1Value = Default, _
	$P2Name = Default, $P2Value = Default, _
	$P3Name = Default, $P3Value = Default, _
	$P4Name = Default, $P4Value = Default, _
	$P5Name = Default, $P5Value = Default, _
	$P6Name = Default, $P6Value = Default, _
	$P7Name = Default, $P7Value = Default, _
	$P8Name = Default, $P8Value = Default, _
	$P9Name = Default, $P9Value = Default  _
	)
	Local Const $ThisFunc = "hr_Log_Msg"
	Local Const $FieldSeperator = Chr(9)
	Local Const $LogFileSpec = "C:\JUNK\HotRod_Log.tsv"

	; Start the Msg string
	Local $MsgStrg = "Script: '" & @ScriptFullPath & "'"

	; If $FuncName is defined - add it.
	If NOT ( ( $FuncName = Default ) OR ( $FuncName = "")  ) Then
		$MsgStrg = $MsgStrg & $FieldSeperator & "Func: '" & $FuncName & "'"
	EndIf

	; If $StepName is defined - add it.
	If NOT ( ( $StepName = Default ) OR ( $StepName = "")  ) Then
		$MsgStrg = $MsgStrg & $FieldSeperator & "Step: '" & $StepName & "'"
	EndIf

	; If $MsgText is defined - add it.
	If NOT ( ( $MsgText = Default ) OR ( $MsgText = "")  ) Then
		$MsgStrg = $MsgStrg & $FieldSeperator & "Message: " & $MsgText
	EndIf

	; If $ErrorValue is defined - add it to message
	IF  Not ( ( $ErrorValue = Default ) OR ( $ErrorValue = "")  OR ( Int( $ErrorValue ) = 0 ) ) Then
		$MsgStrg = $MsgStrg & $FieldSeperator & "Error Value: '" & $ErrorValue & "'"
	EndIf

	; Only include Param that aren't default

		;	2023-09-10 - P1 ONLY can be just the $P1Name. For when the caller has
		;	built the whole shmeer already.
		;
		If $P1Name  <> Default Then $MsgStrg = $MsgStrg & $FieldSeperator & $P1Name
		If $P1Value <> Default Then $MsgStrg = $MsgStrg & ": '" & $P1Value & "'"

	If $P2Name <> Default Then $MsgStrg = $MsgStrg & $FieldSeperator & $P2Name & ": '" & $P2Value & "'"
	If $P3Name <> Default Then $MsgStrg = $MsgStrg & $FieldSeperator & $P3Name & ": '" & $P3Value & "'"
	If $P4Name <> Default Then $MsgStrg = $MsgStrg & $FieldSeperator & $P4Name & ": '" & $P4Value & "'"
	If $P5Name <> Default Then $MsgStrg = $MsgStrg & $FieldSeperator & $P5Name & ": '" & $P5Value & "'"
	If $P6Name <> Default Then $MsgStrg = $MsgStrg & $FieldSeperator & $P6Name & ": '" & $P6Value & "'"
	If $P7Name <> Default Then $MsgStrg = $MsgStrg & $FieldSeperator & $P7Name & ": '" & $P7Value & "'"
	If $P8Name <> Default Then $MsgStrg = $MsgStrg & $FieldSeperator & $P8Name & ": '" & $P8Value & "'"
	If $P9Name <> Default Then $MsgStrg = $MsgStrg & $FieldSeperator & $P9Name & ": '" & $P9Value & "'"

	;	Prefix $MsgStrg with a time stamp

	Local $TimeStamp = @YEAR & "-" & @MON & "-" & @MDAY & "_" & @HOUR & @MIN & @SEC & @MSEC
	$MsgStrg = $TimeStamp & $FieldSeperator & $MsgStrg

	; Write $MsgStrg to the Log file

	If Not FileWriteLine( $LogFileSpec, $MsgStrg ) Then hr_Error_Exit( $ThisFunc, "FileWriteLine( $LogFileSpec, $MsgStrg )", Default, "HotRod Log FileWriteLine failed.", "$LogFileSpec", $LogFileSpec )

EndFunc

