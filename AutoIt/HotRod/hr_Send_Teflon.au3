#include-once

#include "HotRod\hr_Send_Fast.au3"
#include "hr_Keyboard_ModifiersUp.au3"

; =====================================================================
;	Teflon Send
;
;		If my script is triggered by a Ctrl/Alt/Shift hotkey then if any of the
;		modifier keys are still down when I start to SEND, Windows sees it as
;		a hotkey combination. So we wait until the Shift Alt and Ctrl keys are released.
;
; =====================================================================
	Func hr_Send_Teflon( $SendStrg, $SendFlag = 0 )
	Local Const $ThisFunc = "hr_Send_Teflon"

		hr_Keyboard_ModifiersUp( )
		Send( $SendStrg, $SendFlag )

		Return

	EndFunc

; ---------------------------------------------------------------------
;	Fast version of the same thing
; ---------------------------------------------------------------------

	Func hr_Send_TeflonFast( $SendStrg, $SendFlag = 0 )
	Local Const $ThisFunc = "hr_Send_TeflonFast"

		hr_Keyboard_ModifiersUp( )
		hr_Send_Fast( $SendStrg, $SendFlag )

		Return

	EndFunc
