#include-once

#include <Misc.au3>

; ---------------------------------------------------------------------
; 2022-08-26
;
;	Prompt and Pause an AutoIt Console program
;
;	From https://www.autoitscript.com/forum/topic/188161-external-console-programs-ran-by-autoit-are-non-interactive/?do=findComment&comment=1351506
;
; ---------------------------------------------------------------------
Func hr_Console_Pause()
Local Const $ThisFunc = "hr_Console_Pause"

	ConsoleWrite('Press "Enter" to exit' & @CRLF)
	hr_Console_Pause_Wait()

EndFunc

Func hr_Console_Pause_Wait()
	Local $hDLL = DllOpen("user32.dll")
	Local $myWhdl = HWnd(hr_Console_Pause_HandleByPID(@AutoItPID)) ; get Handle of this executable
    Do
        Sleep(50) ; just wait
    Until _IsPressed("0D", $hDLL) And WinActive($myWhdl) ; ENTER key detected
EndFunc

Func hr_Console_Pause_HandleByPID($vProc)
    Local $aWL = WinList()
    For $iCC = 1 To $aWL[0][0]
        If WinGetProcess($aWL[$iCC][1]) = $vProc And _ ; (window with same pid and visible)
                BitAND(WinGetState($aWL[$iCC][1]), 2) Then
            Return $aWL[$iCC][1]
        EndIf
    Next
    Return SetError(2, 0, 0);No windows found
EndFunc