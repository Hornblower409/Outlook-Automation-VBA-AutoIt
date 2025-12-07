#include-once

#Include <Misc.au3>
#include "hr_Keyboard_ModifiersUp.au3"

; =====================================================================
; 2024-11-24 - Use hr_Keyboard_ModifiersUp.au3 instead
; =====================================================================
Func hr_Wait_Teflon( $TimeoutMS = 3000 )
Local Const $ThisFunc = "hr_Wait_Teflon"

	hr_Keyboard_ModifiersUp( Int($TimeoutMS / 1000 ) )

EndFunc
