#include-once

; =====================================================================
;	Wait N secs for a Window to exist, or do an Abort, Retry, Ignore
;
;	Title				<-	Window Title Search String
;	Text				<-	Window Text or ""
;	Time Out (Secs)		<-	Window Exist time out (secs)
;	hWnd				->	Handle of the Window (if found)
;	Return				->	"OK" (if found), "Abort" (User choice), "Ignore" (User choice)

; =====================================================================
Func hr_Window_ExistOrARI( $WinTitle, $WinText, $TimeOutSecs, ByRef $WinHnd, ByRef $Rtn )
Local Const $ThisFunc = "hr_Window_ExistOrARI"

	Local $Retry
	Local $Answer

	Do

		$Retry = False
		$Rtn = "OK"

		; Wait for the Window to exist or timeout
		$WinHnd = WinWait( $WinTitle, $WinText, $TimeOutSecs )
		If $WinHnd <> 0 Then ExitLoop

		; If timeout - ask what to do
		$Answer = MsgBox(18,"Window Exist Failed","Script = '" & @ScriptFullPath & "'." & @CRLF & "Func = '" & $ThisFunc & "'." & @CRLF & @CRLF & "Window with Title = '" & $WinTitle & "', did not exist after '" & $TimeOutSecs & "' secs.")
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

