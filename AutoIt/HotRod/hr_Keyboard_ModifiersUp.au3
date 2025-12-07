#include-once

#include <Misc.au3>
#include "HotRod\hr_Error_Exit.au3"

;	Wait for all Modifier Keys Released
;
;		If my script is triggered by a Ctrl/Alt/Shift/Win hotkey then if any of the
;		modifier keys are still down when I start to SEND, receiver sees it as
;		a hotkey combination. So we wait until all the Modifiers are released.
;
;		10 SHIFT key
; 		11 CTRL key
; 		12 ALT key
; 		5B Left Windows key
; 		5C Right Windows key
;
Func hr_Keyboard_ModifiersUp( $TimeOut = 2, $Sleep = 10 )
Local $ThisFunc = "hr_Keyboard_ModifiersUp"

	Do
		Local $TimerInit = TimerInit()
		Do
			If Not (_IsPressed("10") Or _IsPressed("11") Or _IsPressed("12") Or _IsPressed("5B") Or _IsPressed("5C")) Then Return
			Sleep( $Sleep )
		Until ( TimerDiff( $TimerInit ) >= ($TimeOut * 1000) )
		MsgBox( 0, $ThisFunc, "Let go of the Modifier Keys dummy!" )
	Until False

EndFunc