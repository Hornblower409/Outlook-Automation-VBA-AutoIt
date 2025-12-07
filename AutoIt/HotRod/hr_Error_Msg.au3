#include-once

#include <MsgBoxConstants.au3>

; =====================================================================
; 2023-09-23 - hr_Error_Msg - Show Error message
;
;	Usage:
;
;	hr_Error_Msg( $ThisFunc, "StepName", @error, "MsgText", "$P1Name" )
;	hr_Error_Msg( $ThisFunc, "StepName", @error, "MsgText", "$P1Name", $P1Value, ... )
;
; =====================================================================
Func hr_Error_Msg( _
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

	; Start the Msg string
	Local $MsgStrg = "Script: '" & @ScriptFullPath & "'"

	; If $FuncName is defined - add it.
	If NOT ( ( $FuncName = Default ) OR ( $FuncName = "")  ) Then
		$MsgStrg = $MsgStrg & @CRLF & "Func: '" & $FuncName & "'"
	EndIf

	; If $StepName is defined - add it.
	If NOT ( ( $StepName = Default ) OR ( $StepName = "")  ) Then
		$MsgStrg = $MsgStrg & @CRLF & "Step: '" & $StepName & "'"
	EndIf

	; If $MsgText is defined - add it.
	If NOT ( ( $MsgText = Default ) OR ( $MsgText = "")  ) Then
		$MsgStrg = $MsgStrg & @CRLF & @CRLF & "Error: " & $MsgText
	EndIf

	; If $ErrorValue is defined - add it to message
	IF  Not ( ( $ErrorValue = Default ) OR ( $ErrorValue = "")  OR ( Int( $ErrorValue ) = 0 ) ) Then
		$MsgStrg = $MsgStrg & @CRLF & @CRLF & "Error Value: '" & $ErrorValue & "'"
	EndIf

	; Only include Param that aren't default

		;	2023-09-10 - P1 ONLY can be just the $P1Name. For when the caller has
		;	built the whole shmeer already.
		;
		If $P1Name  <> Default Then $MsgStrg = $MsgStrg & @CRLF & @CRLF & $P1Name
		If $P1Value <> Default Then $MsgStrg = $MsgStrg & ": '" & $P1Value & "'"

	If $P2Name <> Default Then $MsgStrg = $MsgStrg & @CRLF & $P2Name & ": '" & $P2Value & "'"
	If $P3Name <> Default Then $MsgStrg = $MsgStrg & @CRLF & $P3Name & ": '" & $P3Value & "'"
	If $P4Name <> Default Then $MsgStrg = $MsgStrg & @CRLF & $P4Name & ": '" & $P4Value & "'"
	If $P5Name <> Default Then $MsgStrg = $MsgStrg & @CRLF & $P5Name & ": '" & $P5Value & "'"
	If $P6Name <> Default Then $MsgStrg = $MsgStrg & @CRLF & $P6Name & ": '" & $P6Value & "'"
	If $P7Name <> Default Then $MsgStrg = $MsgStrg & @CRLF & $P7Name & ": '" & $P7Value & "'"
	If $P8Name <> Default Then $MsgStrg = $MsgStrg & @CRLF & $P8Name & ": '" & $P8Value & "'"
	If $P9Name <> Default Then $MsgStrg = $MsgStrg & @CRLF & $P9Name & ": '" & $P9Value & "'"

	; Show the box

	MsgBox( _
		$MB_ICONERROR, _
		@ScriptName, _
		$MsgStrg _
		)

EndFunc

