; ---------------------------------------------------------------------
;	!!!!    Don't use for new work. Use hr_Registry.au3 instead  !!!!
; ---------------------------------------------------------------------

#include-once

#include "HotRod\hr_Script_Error_Exit.au3"

; =====================================================================
;	Simple Registry Functions
; =====================================================================

	Func hr_RegRead( $RegKey, $RegName )
	Local Const $ThisFunc = "hr_RegRead"

		Local $RegValue = RegRead( $RegKey, $RegName )
		If @error Then
			hr_Script_Error_Exit( $ThisFunc, 'Reg Read failed.', @error, _
				"RegKey", $RegKey, _
				"RegName", $RegName _
			)
		EndIf

		Return $RegValue

	EndFunc

	Func hr_RegWrite( $RegKey, $RegName, $RegType, $RegValue )
	Local Const $ThisFunc = "hr_RegWrite"

		RegWrite( $RegKey, $RegName, $RegType, $RegValue )
		If @error Then
			hr_Script_Error_Exit( $ThisFunc, 'Reg Write failed.', @error, _
				"RegKey", $RegKey, _
				"RegName", $RegName, _
				"RegType", $RegType _
			)
		EndIf

	EndFunc

