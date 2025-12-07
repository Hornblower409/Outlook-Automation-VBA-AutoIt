#include-once

; =====================================================================
;	Send Fast - Decrease the delay between keystrokes and send.
; =====================================================================

	Func hr_Send_Fast( $SendStrg, $SendFlag = 0 )

		AutoItSetOption ( "SendKeyDelay", 0 )
		AutoItSetOption ( "SendKeyDownDelay",0 )

		Send( $SendStrg, $SendFlag )

		; Reset to the default.

		AutoItSetOption ( "SendKeyDelay", Default )
		AutoItSetOption ( "SendKeyDownDelay", Default )

		Return

	EndFunc
