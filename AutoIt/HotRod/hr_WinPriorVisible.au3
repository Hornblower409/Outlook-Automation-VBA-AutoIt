#include-once

#include <WinAPISysWin.au3>
#include <WinAPI.au3>

; =====================================================================
; 2022-09-29
;
;	Find the Previously Active Visible Window
;	Assume that this script is currently Active
;
; =====================================================================
Func hr_WinPriorVisible()
Local Const $ThisFunc = "hr_WinPriorVisible"

	;	Start with my handle
	;
	Local $hWnd = WinGetHandle( "[Active]" )
	If $hWnd = 0 Then hr_Error_Exit( $ThisFunc, 'WinGetHandle( "[Active]" )', Default, "Failed" )

	Do

		;	Get the handle of the previous Window in the Z order
		;	(Seems like it shoud say PRIOR, not NEXT. But that's how it works)
		;
		$hWnd = _WinAPI_GetWindow($hWnd, $GW_HWNDNEXT)
		If $hWnd = 0 Then hr_Error_Exit( $ThisFunc, "_WinAPI_GetWindow($hWnd, $GW_HWNDNEXT)", Default, "Failed", "$hWnd", $hWnd )

		;	Get this window's State
		Local $winState = WinGetState( $hWnd )
		If @error Then hr_Error_Exit( $ThisFunc, "WinGetState( $hWnd )", @error, "Failed", "$hWnd", $hWnd )

	; If Exist and Visible and Enabled and NOT minimized - we'll take it
	Until _
		        BitAND( $winState, $WIN_STATE_EXISTS ) _
		And     BitAND( $winState, $WIN_STATE_VISIBLE ) _
		And     BitAND( $winState, $WIN_STATE_ENABLED ) _
		And Not BitAND( $winState, $WIN_STATE_MINIMIZED )

	;	Return the handle
	;
	Return $hWnd

EndFunc