#include "HotRod\hr_Directives.au3"
#include "HotRod\hr_Error_Exit.au3"
#include "HotRod\hr_Log_Msg.au3"

#include "HotRod\hr_IsWorkstationLocked.au3"

#include "Outlook\ProcXeq.au3"

#NoTrayIcon

; =====================================================================
; 2023-12-02
;
;	Clear any "Your IMAP server closed the connection" messages.
;	Execute the UpdateFolder proc via ProcXeq.
;
;	2025-03-09 - Run as Scheduled Task Outlook_UpdateFolder
;
;	SPOS - The "Your IMAP server closed the connection" message is modal.
;	Stops my ProcXeq from running. So we have to clear any
;	message first and then run the UpdateFolder proc.
;
; =====================================================================
UpdateFolder()
Func UpdateFolder()
Local Const $ThisFunc = "UpdateFolder"

;~ 	hr_Log_Msg( $ThisFunc, "Initial Entry" )

	;	Check Command Line Arg
	;
	If $CmdLine[0] <> 1 Then hr_Error_Exit( $ThisFunc, "Check Command Line Arg", Default, "Command Line must have a single arg. The KnownPath to the IMAP folder.")
	Local $KnownPath = $CmdLine[1]

	;	If Workstation is locked WinActivate will always fail - So silent exit.
	;
	If hr_IsWorkstationLocked() Then Exit( 0 )

	;	Clear any "Your IMAP server closed the connection" message.
	;
	Local $winTitle = "[TITLE:Microsoft Outlook; CLASS:#32770]"
	Local $winText = "Your IMAP server closed the connection."
	If WinExists( $winTitle, $winText) Then
		Local $hWnd = WinActivate( $winTitle, $winText )
		If $hWnd = 0 Then hr_Error_Exit( $ThisFunc, "WinActivate IMAP message", Default, "Found but failed to Activate the 'Your IMAP server closed the connection' window.")
		Send( "{ENTER}")
;~ 		hr_Log_Msg( $ThisFunc, "Cleared Your IMAP server closed the connection message." )
	EndIf

	;	Execute the UpdateFolder Proc
	;
;~ 	hr_Log_Msg( $ThisFunc, "Calling Outlook_ProcXeq( UpdateFolder )" )
;
;	If AutoIt had Named Args:
;	Outlook_ProcXeq( CmdLine:="UpdateFolder $KnownPath", $Option_OutlookWindowActive:=False, $Option_SilentErrorExit:=True )
;
	Local $PXCommandLine[3] = [2, "UpdateFolder", $KnownPath]
	Outlook_ProcXeq( $PXCommandLine, False, True )

;~ 	hr_Log_Msg( $ThisFunc, "Returned from Outlook_ProcXeq( UpdateFolder )" )

	Exit( 0 )

EndFunc
