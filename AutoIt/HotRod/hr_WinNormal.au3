#include-once

#include <WinAPISysWin.au3>
#include <WinAPI.au3>

; =====================================================================
; 2022-09-29
;
;	Test if a window is "Normal" (Exist, Visible, Enabled, Not Minimized)
;
;		->	Handle to the window to test.
;		<-	True if Normal Else False.
;
; =====================================================================
Func hr_WinNormal( $hWnd ) ; As Boolean
Local Const $ThisFunc = "hr_WinNormal"

	;	Get the window's State
	Local $winState = WinGetState( $hWnd )
	If @error Then hr_Error_Exit( $ThisFunc, "WinGetState( $hWnd )", @error, "Failed", "$hWnd", $hWnd )

	;	If Exist and Visible and Enabled and NOT minimized - we'll take it
	Return _
		        BitAND( $winState, $WIN_STATE_EXISTS ) _
		And     BitAND( $winState, $WIN_STATE_VISIBLE ) _
		And     BitAND( $winState, $WIN_STATE_ENABLED ) _
		And Not BitAND( $winState, $WIN_STATE_MINIMIZED )

EndFunc