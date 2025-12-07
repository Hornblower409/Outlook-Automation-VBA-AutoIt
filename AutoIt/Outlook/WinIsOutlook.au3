#include-once

#include <Process.au3>
#include <WinAPISysWin.au3>

; =====================================================================
; 2022-09-30
;
;	Is a window an Outlook window?
;
;	->	$hWnd
;	<-	True - Is an Outlook Window. Else False.
;
; =====================================================================
Func Outlook_WinIsOutlook( $hWnd )
Local Const $ThisFunc = "Outlook_WinIsOutlook"

	;	If Class <> "rctrl_renwnd32" then not an Outlook Window
	;
	Local $Class = _WinAPI_GetClassName( $hWnd )
	If $Class = "" Then hr_Error_Exit( $ThisFunc, "_WinAPI_GetClassName", Default, "Failed" )
	If $Class <> "rctrl_renwnd32" Then Return False

	;	Get the PID and ProcessName for the window
	;
	Local $PID = WinGetProcess( $hWnd )
	If $PID = -1 Then hr_Error_Exit( $ThisFunc, "WinGetProcess", Default, "Failed" )
	Local $PName = _ProcessGetName( $PID )
	If $PName = ""  Then hr_Error_Exit( $ThisFunc, "_ProcessGetName", Default, "Failed" )

	;	If Process Name <> "Outlook.exe" then not an Outlook window
	;
	If StringUpper( $PName ) <> "OUTLOOK.EXE" Then Return False

	;	Must be
	;
	Return True

EndFunc
