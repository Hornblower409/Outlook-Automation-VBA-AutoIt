#include-once
#include "HotRod\hr_Error_Exit.au3"
#include <StringConstants.au3>

; =====================================================================
; 2022-10-03
;
;	COM Error Exit
;
;	Usage
;
;		#include "HotRod\hr_COM_Error_Exit.au3"
;
;		In the MAIN script ONLY - You can override the Globals by creating a Local of the same name.
;		The Local will remain in force until you exit the script or enter a function.
;
;			Local $hr_COM_Error_OnError[1] = [$ThisFunc]
;			Local $hr_COM_Error_OnError[2] = [$ThisFunc, "Step"]
;			Local $hr_COM_Error_OnError[3] = [$ThisFunc, "Step", "Message"]
;			Local $hr_COM_Error_OnError = ""  {No message box. Just Exit with Error}
;			Local $hr_COM_Error_Ignore[n] = ["0x8000404", "0x{HexError}", ... ]
;			Local $hr_COM_Error_Exit = False
;
;		When $hr_COM_Error_Exit = False you must check for errors after every COM operation.
;
;			$hr_COM_Error_Exit = False
;			{Do some com}
;			If $hr_COM_Error <> 0 Then ...
;			$hr_COM_Error = 0
;
; =====================================================================

Global $hr_COM_Error_Exit_Handler = ObjEvent("AutoIt.Error", "hr_COM_Error_Exit")
Global $hr_COM_Error_OnError = ""
Global $hr_COM_Error_Ignore = ""
Global $hr_COM_Error_Exit = True
Global $hr_COM_Error = 0

Func hr_COM_Error_Exit( $oError )
Local Const $ThisFunc = "hr_COM_Error_Exit"

	If IsArray( $hr_COM_Error_Ignore ) Then
		For $HexErr In $hr_COM_Error_Ignore
			If Number($HexErr) = Number("0x" & Hex($oError.number)) Then Return
		Next ; $HexErr
	EndIf

	If Not $hr_COM_Error_Exit Then
		$hr_COM_Error = $oError.number
		Return
	EndIf

	Local $ComFunc = $ThisFunc
	Local $ComStep = "Global COM Error Trap"
	Local $ComMsg = "COM Error"
	If IsArray( $hr_COM_Error_OnError ) Then
		If UBound( $hr_COM_Error_OnError ) > 0 Then $ComFunc = $hr_COM_Error_OnError[0]
		If UBound( $hr_COM_Error_OnError ) > 1 Then $ComStep = $hr_COM_Error_OnError[1]
		If UBound( $hr_COM_Error_OnError ) > 2 Then $ComMsg = $hr_COM_Error_OnError[2]
	Else
;		Exit( $oError.number )
	EndIf

	hr_Error_Exit( _
		$ComFunc, _
		$ComStep, _
		"0x" & Hex($oError.number), _
		$ComMsg, _
		"win description", StringStripWS ( $oError.windescription, $STR_STRIPLEADING + $STR_STRIPTRAILING ), _
		"description", $oError.description, _
		"source", $oError.source, _
		"last dllerror", $oError.lastdllerror, _
		"script line", $oError.scriptline, _
		"retcode", "0x" & Hex($oError.retcode) _
	)

EndFunc

