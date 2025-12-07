#include-once
#include <Misc.au3>

; =====================================================================
;	Example Calls:
;
;	#include "HotRod\hr_Debug_MsgBox.au3"
;
;	hr_Debug_MsgBox( $ThisFunc, "StepName", "MsgText", "$P1Name", $P1Value, ... )
;
;	hr_Debug_MsgBox( $ThisFunc, "StepName", _
;		"MsgText", _
;		"$P1Name", $P1Value, ... )
;
;	If reply to the MsgBox is Cancel - Exits the script.
; =====================================================================
Func hr_Debug_MsgBox( _
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

	; If $MsgText is defined - add it
	If Not ( ( $MsgText = Default ) OR ( $MsgText = "")  ) Then
		$MsgStrg = $MsgStrg & @CRLF & @CRLF & "Message: '" & $MsgText & "'"
	EndIf

	; Only include Param that aren't default
	If $P1Name <> Default Then $MsgStrg = $MsgStrg & @CRLF & @CRLF & $P1Name & ": '" & $P1Value & "'"
	If $P2Name <> Default Then $MsgStrg = $MsgStrg & @CRLF & @CRLF & $P2Name & ": '" & $P2Value & "'"
	If $P3Name <> Default Then $MsgStrg = $MsgStrg & @CRLF & @CRLF & $P3Name & ": '" & $P3Value & "'"
	If $P4Name <> Default Then $MsgStrg = $MsgStrg & @CRLF & @CRLF & $P4Name & ": '" & $P4Value & "'"
	If $P5Name <> Default Then $MsgStrg = $MsgStrg & @CRLF & @CRLF & $P5Name & ": '" & $P5Value & "'"
	If $P6Name <> Default Then $MsgStrg = $MsgStrg & @CRLF & @CRLF & $P6Name & ": '" & $P6Value & "'"
	If $P7Name <> Default Then $MsgStrg = $MsgStrg & @CRLF & @CRLF & $P7Name & ": '" & $P7Value & "'"
	If $P8Name <> Default Then $MsgStrg = $MsgStrg & @CRLF & @CRLF & $P8Name & ": '" & $P8Value & "'"
	If $P9Name <> Default Then $MsgStrg = $MsgStrg & @CRLF & @CRLF & $P9Name & ": '" & $P9Value & "'"

	; Show the box. Exit if they pick Cancel
	If MsgBox( _
		$MB_ICONINFORMATION + $MB_OKCANCEL, _
		@ScriptName, _
		$MsgStrg _
		) _
	= $IDCANCEL Then Exit( 1 )

EndFunc