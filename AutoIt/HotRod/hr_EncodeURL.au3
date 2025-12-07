#include-once

; ---------------------------------------------------------------------
;	From https://www.autoitscript.com/forum/topic/164408-url-encoder-decoder-functions/
; ---------------------------------------------------------------------

	Func hr_EncodeUrl($src)
		Local $i
		Local $ch
		Local $NewChr
		Local $buff

		;Init Counter
		$i = 1

		While ($i <= StringLen($src))
			;Get byte code from string
			$ch = Asc(StringMid($src, $i, 1))

			;Look for what bytes we have
			Switch $ch
				;Looks ok here
				Case 45, 46, 48 To 57, 65 To 90, 95, 97 To 122, 126
					$buff &= Chr($ch)
					;Space found
				Case 32
					$buff &= "+"
				Case Else
					;Convert $ch to hexidecimal
					$buff &= "%" & Hex($ch, 2)
			EndSwitch
			;INC Counter
			$i += 1
		WEnd

		Return $buff

	EndFunc

; ---------------------------------------------------------------------