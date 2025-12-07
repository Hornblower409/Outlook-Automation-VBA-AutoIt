#include-once

; =====================================================================
;	Expand any Enviornment Variables (%Var%) in a string
; =====================================================================
Func hr_ExpandEnvStrings( $String )
;~ Local Const $ThisFunc = "hr_ExpandEnvStrings"

	AutoItSetOption( "ExpandEnvStrings" , True )

;	Force an Eval. Otherwise he won't do the expansion
		;
		$String = $String & ""

	AutoItSetOption( "ExpandEnvStrings" , Default )

	Return $String

EndFunc

