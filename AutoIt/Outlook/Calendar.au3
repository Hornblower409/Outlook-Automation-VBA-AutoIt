#include-once

; ---------------------------------------------------------------------
;  Calendar 3 Day View from Today
; ---------------------------------------------------------------------
Func calendar_Today3Day( )
Local Const $ThisFunc = "calendar_Today3Day"

	Local Const $SleepMS = 250

	Send( "!vr" )						; View -> Day
	Sleep( $SleepMS )

	Send( "!hod") 						; Go -> Today
	Sleep( $SleepMS )

	Send("^{END}")						; Scroll to end
	Sleep( $SleepMS )

	Send( "!3" )						; Show 3 days
	Sleep( $SleepMS )

EndFunc
