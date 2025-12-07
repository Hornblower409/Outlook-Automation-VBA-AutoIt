#include-once

#include "Outlook\WinIsOutlook.au3"

; ---------------------------------------------------------------------
;	Get the Topmost Outlook Window
; ---------------------------------------------------------------------
Func Outlook_Top()
Local Const $ThisFunc = "Outlook_Top"

	;	Get a Z-Order list if all windows with an Outlook Class
	;	If no matches - throw an error
	;
	Local $wList = WinList( "[CLASS:rctrl_renwnd32]" )
	If $wList[0][0] = 0 Then hr_Error_Exit( $ThisFunc, "Check WinList Count", Default, "No windows found with '[CLASS:rctrl_renwnd32]'." )

	;	Get the topmost matching visable window that is an Outlook Window
	;	If no hit - throw an error
	;
	Local $hFound = 0
	For $i = 1 To $wList[0][0]

		If BitAND(WinGetState($wList[$i][1]), 2) Then
			If Outlook_WinIsOutlook( $wList[$i][1] ) Then
				$hFound = $wList[$i][1]
				ExitLoop
			EndIf
		EndIf

	Next ; $i
	If $hFound = 0 Then hr_Error_Exit( $ThisFunc, "Find Outlook Window", Default, "No visable Outlook window found.")

	Return $hFound

EndFunc

