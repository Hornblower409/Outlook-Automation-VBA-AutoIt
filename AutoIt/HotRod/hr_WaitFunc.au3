#include-once

; =====================================================================
;	Wait for a Function to return non-Zero/non-Empty or timeout
;	Returns the return value from the Func or @error = 0xF0F0 on a timeout
; =====================================================================

Func hr_WaitFunc( $FuncName, $TimeOut, $Sleep = 100 )
	Local $ThisFunc = "hr_WaitFunc"
    Local $secs = TimerInit()

	Do

		Local $Rtn = Call( $FuncName )
		If @error = 0xDEAD And @extended = 0xBEEF Then
			hr_Script_Error_Exit( $ThisFunc, "Function does not exist", @error, "$FuncName", $FuncName )
		EndIf
		If NOT( $Rtn == "0" ) AND NOT( $Rtn == "" ) Then Return SetError( 0, 0, $Rtn )

		Sleep( $Sleep )

    Until TimerDiff($secs) >= $TimeOut * 1000
    Return SetError( 0xF0F0, 0, 0 )

EndFunc

