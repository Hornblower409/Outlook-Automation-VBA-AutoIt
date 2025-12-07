#include-once

#include "HotRod\hr_Error_Exit.au3"

; =====================================================================
; 2022-10-04
;
;	Put text on the clipboard
;
; =====================================================================
Func hr_ClipPutText( $Text )
Local Const $ThisFunc = "hr_ClipPutText"

	Local $Rtn = ClipPut( $Text )
	If $Rtn = 0 Then hr_Error_Exit( $ThisFunc, Default, Default, "ClipPut Failed" )

EndFunc
