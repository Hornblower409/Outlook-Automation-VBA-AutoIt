#include-once

; =====================================================================
; Convert a Raw string to a cooked strin
;
;	$Raw	-> Raw input string
;
;	<-
;			CR and TAB -> {XXX}
;			! # + ^ { } -> {X}
;			Any other Control Chars -> Dropped
;
; =====================================================================

Func hr_Send_StringCook(  $Raw )
Local $ThisFunc = "hr_Send_StringCook"

	Local $Cooked = ""
	For $i = 1 To StringLen( $Raw )

		Local $RawChr = StringMid( $Raw, $i, 1 )
		Switch Asc( $RawChr )

			Case 9   ; TAB
				$Cooked &= "{TAB}"
			Case 13  ; CR
				$Cooked &= "{ENTER}"
			Case 0 TO 31 ; Any other Control Char
				; Drop it

			Case 33  ; !
				$Cooked &= "{!}"
			Case 35  ; #
				$Cooked &= "{#}"
			Case 43  ; +
				$Cooked &= "{+}"
			Case 94  ; ^
				$Cooked &= "{^}"
			Case 123 ; {
				$Cooked &= "{}}"
			Case 125 ; }
				$Cooked &= "{}}"

			Case Else ; All others
				$Cooked &= $RawChr ; As-Is

		EndSwitch

	Next ; $i

	Return $Cooked

EndFunc

