#include-once

; =====================================================================
;	Activate a Window in N secs, or do an Abort, Retry, Ignore
;
;	$WinTitle			<-	Window Title Search String
;	$WinText			<-	Window Text or ""
;	$TimeOutSecs		<-	WinWaitActive time out (secs)
;	$WinHnd				->	Handle of the Window (if found)
;	$Rtn				->	"OK" (if found), "ABORT" (User choice), "IGNORE" (User choice)
; =====================================================================
Func hr_Window_ActivateOrARI( $WinTitle, $WinText, $TimeOutSecs, ByRef $WinHnd, ByRef $Rtn )
Local Const $ThisFunc = "hr_Window_ActivateOrARI"

	Local $Retry
	Local $Answer
	Local $ElapsedSecs = 0

	Do

		$Retry = False
		$Rtn = "OK"

		;	2023-08-01 - Inner Loop
		;
		;		For when the Window opens but doesn't get focus.
		;		Keep trying an Activate/WaitActivate every second.
		;
		Do

			; If it was already active or activates immediatley - return OK
			$WinHnd = WinActivate( $WinTitle, $WinText )
			If $WinHnd <> 0 Then Return

			; If it Activates in One Sec - return OK
			$WinHnd = WinWaitActive( $WinTitle , $WinText, 1 )
			If $WinHnd <> 0 Then Return

			; If not timeout - try again
			$ElapsedSecs = $ElapsedSecs + 1

		Until $ElapsedSecs > $TimeOutSecs

		; Timeout  - ask what to do
		$Answer = MsgBox(18,"Window Activate Failed","Script = '" & @ScriptFullPath & "'." & @CRLF & "Func = '" & $ThisFunc & "'." & @CRLF & @CRLF & "Window with Title = '" & $WinTitle & "', did not Activate after '" & $TimeOutSecs & "' secs.")
		Select
			Case $Answer = 3 ; Abort
				$Rtn = "ABORT"
			Case $Answer = 4 ; Retry
				$Retry = True
			Case $Answer = 5 ; Ignore
				$Rtn = "IGNORE"
		EndSelect

	Until $Retry = False

EndFunc
