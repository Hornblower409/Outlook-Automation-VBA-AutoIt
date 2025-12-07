#include-once

; =====================================================================
;
;	#include "HotRod\hr_Debug_Print.au3"
;
;	Example Calls:
;
;	hr_Debug_Print( $ThisFunc, "StepName", "MsgText", "$P1Name", $P1Value, ... )
;	hr_Debug_Print( $ThisFunc, "StepName", Default, "$P1Name", $P1Value, ... )
;	hr_Debug_Print( $ThisFunc, Default, Default, "$P1Name", $P1Value, ... )
;
;	hr_Debug_Print( $ThisFunc, "StepName", _
;		"MsgText", _
;		"$P1Name", $P1Value, ... )
;
;
;	!!!!!   Will NOT work when called from ascript running Elevated in Scite !!!!!
;
;	See https://www.autoitscript.com/forum/topic/122348-how-to-consolewrite-after-requireadmin/?do=findComment&comment=849338
;
; =====================================================================

Func hr_Debug_Print( _
	Const $FuncName = Default, _
	$StepName = Default, _
	$MsgText = Default, _
	$P1Name = Default, $P1Value = Default, _
	$P2Name = Default, $P2Value = Default, _
	$P3Name = Default, $P3Value = Default, _
	$P4Name = Default, $P4Value = Default, _
	$P5Name = Default, $P5Value = Default, _
	$P6Name = Default, $P6Value = Default, _
	$P7Name = Default, $P7Value = Default, _
	$P8Name = Default, $P8Value = Default, _
	$P9Name = Default, $P9Value = Default _
	)

	Local $MsgStrg = "DEBUG_PRINT"

	If NOT ( ( $FuncName = Default ) OR ( $FuncName = "")  ) Then
		$MsgStrg = $MsgStrg & " | " & $FuncName
	EndIf

	If NOT ( ( $StepName = Default ) OR ( $StepName = "")  ) Then
		$MsgStrg = $MsgStrg & " | " & $StepName
	EndIf

	If NOT ( ( $MsgText = Default ) OR ( $MsgText = "")  ) Then
		$MsgStrg = $MsgStrg & " | " & $MsgText
	EndIf

	If $P1Name <> Default Then $MsgStrg = $MsgStrg & " | " & $P1Name & ": '" & $P1Value & "'"
	If $P2Name <> Default Then $MsgStrg = $MsgStrg & " | " & $P2Name & ": '" & $P2Value & "'"
	If $P3Name <> Default Then $MsgStrg = $MsgStrg & " | " & $P3Name & ": '" & $P3Value & "'"
	If $P4Name <> Default Then $MsgStrg = $MsgStrg & " | " & $P4Name & ": '" & $P4Value & "'"
	If $P5Name <> Default Then $MsgStrg = $MsgStrg & " | " & $P5Name & ": '" & $P5Value & "'"
	If $P6Name <> Default Then $MsgStrg = $MsgStrg & " | " & $P6Name & ": '" & $P6Value & "'"
	If $P7Name <> Default Then $MsgStrg = $MsgStrg & " | " & $P7Name & ": '" & $P7Value & "'"
	If $P8Name <> Default Then $MsgStrg = $MsgStrg & " | " & $P8Name & ": '" & $P8Value & "'"
	If $P9Name <> Default Then $MsgStrg = $MsgStrg & " | " & $P9Name & ": '" & $P9Value & "'"

	ConsoleWrite( $MsgStrg & @CRLF )

EndFunc