#include-once

#include <AutoItConstants.au3>

; =====================================================================
; 2022-09-25
;
;	Wait for any window matching an AWD to have STATE.
;	Returns hWnd of a matching window or Zero if timed out.
;
; =====================================================================
Func hr_WinWaitState( $AWD, $WinState, $TimeOut, $Sleep = 250 )
;~ Local Const $ThisFunc = "hr_WinWaitState"

	Local $TimerInit = TimerInit()

	Do

		Local $WinList = WinList( $AWD )
		For $i = 1 To $WinList[0][0]
;~ 			ConsoleWrite( $WinList[$i][0] & " | " & WinGetState( $WinList[$i][1] ) & " | " & $WinList[$i][1] & $NewLine )
			If WinGetState( $WinList[$i][1] ) = $WinState Then
				Return $WinList[$i][1]
			EndIf
		Next

		Sleep( $Sleep )

	Until ( TimerDiff( $TimerInit ) >= ($TimeOut * 1000) )

	Return 0

EndFunc
