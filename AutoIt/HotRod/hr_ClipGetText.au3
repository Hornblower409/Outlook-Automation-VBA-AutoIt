#include-once

#include "HotRod\hr_Error_Exit.au3"

; =====================================================================
; 2022-10-04
;
;	Get the clipboard text or error exit
;
;	Often (always?) there is a delay after you Send("^c") before
;	the Clipboard is updated. No error and no indication that what you
;	are getting is stale. Hence the optional $Delay param. 0 = No Delay.
;
; =====================================================================
Func hr_ClipGetText( $Caller = Default, $Delay = 50)
Local Const $ThisFunc = "hr_ClipGetText"

	If $Caller = Default Then $Caller = $ThisFunc
	Sleep( $Delay )

	Local $Rtn = ClipGet( )
	Switch @error
		Case 0
			If $Rtn = "" Then hr_Error_Exit( $Caller, "ClipGet", Default, "Clipboard entry is an empty string" )
			Return $Rtn
		Case 1
			hr_Error_Exit( $Caller, "ClipGet", @error, "Clipboard is empty" )
		Case 2
			hr_Error_Exit( $Caller, "ClipGet", @error, "Clipboard entry is not text" )
		Case 3, 4
			hr_Error_Exit( $Caller, "ClipGet", @error, "Clipboard is not accessable" )

	EndSwitch

EndFunc
