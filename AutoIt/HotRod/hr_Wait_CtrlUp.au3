
; ---------------------------------------------------------------------
; 	2024-11-24 - Use "HotRod\hr_Keyboard_ModifiersUp.au3" instead
;	for all new work.
; ---------------------------------------------------------------------

#include-once

#include "HotRod\hr_Keyboard_ModifiersUp.au3"

; ---------------------------------------------------------------------
; 2023-08-19
;
;	Wait for Ctrl Keys Up
;
; ---------------------------------------------------------------------
Func hr_Wait_CtrlUp()
Local Const $ThisFunc = "hr_Wait_CtrlUp"

	hr_Keyboard_ModifiersUp( )

EndFunc
