#include-once

; =====================================================================
;	Stupid function for Stupid HPSM
;
;	Because he has to have time to highlight the element under the
;	mouse before the Click will work and it takes a SLOW click to register
; =====================================================================
Func hr_MouseMoveWaitClick( $x, $y, $Sleep = 1000 )
Local $ThisFunc = "hr_MouseMoveWaitClick"

	Opt( "MouseClickDownDelay", 20 )
	MouseMove( $x, $y, 0 )
	Sleep( $Sleep )
	MouseClick( $MOUSE_CLICK_PRIMARY )
	Opt( "MouseClickDownDelay", Default )

EndFunc

