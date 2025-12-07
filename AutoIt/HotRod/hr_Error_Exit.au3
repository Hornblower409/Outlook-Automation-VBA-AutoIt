#include-once

#include "HotRod\hr_Error_Msg.au3"

; =====================================================================
; 2023-09-23 - hr_Error_Exit - Show Error message and exit.
;
;	Usage:
;
;	hr_Error_Exit( $ThisFunc, "StepName", @error, "MsgText", "$P1Name" )
;	hr_Error_Exit( $ThisFunc, "StepName", @error, "MsgText", "$P1Name", $P1Value, ... )
;
; =====================================================================
Func hr_Error_Exit( _
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

	hr_Error_Msg( $FuncName, $StepName, $ErrorValue, $MsgText, $P1Name, $P1Value, $P2Name, $P2Value, $P3Name, $P3Value, $P4Name, $P4Value, $P5Name, $P5Value, $P6Name, $P6Value, $P7Name, $P7Value, $P8Name, $P8Value, $P9Name, $P9Value )

	; Default ExitCode to one.
	; If $ErrorValue is defined - use it for the Exit Code
	;
	Local $Exit = 1
	IF  Not ( ( $ErrorValue = Default ) OR ( $ErrorValue = "")  OR ( Int( $ErrorValue ) = 0 ) ) Then
		$Exit = $ErrorValue
	EndIf

	Exit( Int( $Exit ) )

EndFunc

;~ ; 2023-09-23 - What was this doing here?
;~ ;
;~ Func hr_Error_Exit_IsHex($sString)
;~     Return StringRegExp($sString, '^0x[[:xdigit:]]+$') > 0 ; Or the class [0-9A-Za-z].
;~ EndFunc   ;==>_IsHex
