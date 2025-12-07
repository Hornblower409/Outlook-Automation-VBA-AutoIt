; =====================================================================
;	!!!!  DO NOT USE THIS ONE IN NEW CODE  !!!!
;	!!!!  USE hr_Error_Exit				   !!!!
; =====================================================================
#include-once
#include <Misc.au3>
#include "HotRod\hr_Error_Exit.au3"

Func hr_Script_Error_Exit( Const $FuncName, $ErrorDesc, $ErrorValue = Default, $P1A = Default, $P1B = Default, $P2A = Default, $P2B = Default, $P3A = Default, $P3B = Default )

;	Pass the call to hr_Error_Exit
;
hr_Error_Exit( _
	$FuncName, _
	Default, _
	$ErrorValue, _
	$ErrorDesc, _
	$P1A, $P1B, _
	$P2A, $P2B, _
	$P3A, $P3B _
	)

;~Local Const $ThisFunc = "hr_Script_Error_Exit"
;~
;~	Local $MsgStrg = ""
;~	Local $Exit = 1
;~
;~	; If $ErrorValue is defined - add a line in the dialog and use it for the Exit Code
;~
;~	If NOT ( ( $ErrorValue = Default ) OR ( $ErrorValue = "")  ) Then
;~
;~		$MsgStrg = $MsgStrg & "Error Value: '" & $ErrorValue & "'" & @CRLF & @CRLF
;~		$Exit = $ErrorValue
;~
;~	EndIf
;~
;~	; Only include Param that aren't default
;~
;~	If $P1A <> Default Then $MsgStrg = $MsgStrg & $P1A & ": '" & $P1B & "'" & @CRLF & @CRLF
;~	If $P2A <> Default Then $MsgStrg = $MsgStrg & $P2A & ": '" & $P2B & "'" & @CRLF & @CRLF
;~	If $P3A <> Default Then $MsgStrg = $MsgStrg & $P3A & ": '" & $P3B & "'" & @CRLF & @CRLF
;~
;~	; Show the box
;~
;~	MsgBox( _
;~		$MB_ICONERROR, _
;~		@ScriptName, _
;~		"Script: '" & @ScriptFullPath & "'" & @CRLF & @CRLF _
;~		& "Function: '" & $FuncName & "'" & @CRLF & @CRLF _
;~		& "Error Desc: '" & $ErrorDesc & "'"  & @CRLF & @CRLF _
;~		& $MsgStrg _
;~	)
;~
;~	; We're out of here
;~
;~	Exit( Int( $Exit ) )
;~
EndFunc
