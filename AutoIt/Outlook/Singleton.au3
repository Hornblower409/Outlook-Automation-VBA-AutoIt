#include-once

#include <WinAPIProc.au3>
#include <Misc.au3>

#include "HotRod\hr_Error_Exit.au3"

; Singleton (only one instance allowed) Mutex
Global Const $Outlook_Common_Singleton_Name = "HotRod_Outlook"

;	Wait for the Singleton. Return the Singleton Mutex Handle or Exit if timed out.
;
Func Outlook_Singleton_Wait_Exit( $TimeOutMSecs = 5000 )
;~ Local Const $ThisFunc = "Outlook_Singleton_Wait_ErrorExit"

	Local $hMutex = Outlook_Singleton_Wait( $TimeOutMSecs )
	If $hMutex = 0 Then Exit( 1 )
	Return $hMutex

EndFunc

;	Wait for the Singleton. Return the Singleton Mutex Handle or Loud Error Exit if timed out.
;
Func Outlook_Singleton_Wait_ErrorExit( $TimeOutMSecs = 5000 )
Local Const $ThisFunc = "Outlook_Singleton_Wait_ErrorExit"

	Local $hMutex = Outlook_Singleton_Wait( $TimeOutMSecs )
	If $hMutex = 0 Then hr_Error_Exit( $ThisFunc, "Outlook_Singleton_Wait", Default, "Failed to own the Outlook Singleton Mutex after " & $TimeOutMSecs & "ms.")
	Return $hMutex

EndFunc

;	Wait for the Singleton. Return the Singleton Mutex Handle or Zero if timed out.
;
Func Outlook_Singleton_Wait( $TimeOutMSecs = 5000 )
;~ Local Const $ThisFunc = "Outlook_Singleton_Wait"

	Local $WaitedMSecs = 0
	Local Const $WaitCheckMSecs = 100

	Do
		Local $hMutex = _Singleton( $Outlook_Common_Singleton_Name, 1 )
		If $hMutex <> 0 Then Return $hMutex
		Sleep( $WaitCheckMSecs )
		$WaitedMSecs = $WaitedMSecs + $WaitCheckMSecs
	Until $WaitedMSecs > $TimeOutMSecs

	Return 0

EndFunc

; 	Error exit if the Singleton is owned by another process
;
Func Outlook_Singleton_ErrorExit()
Local Const $ThisFunc = "Outlook_Singleton_ErrorExit"

	If _Singleton( $Outlook_Common_Singleton_Name, 1 ) = 0 Then
		hr_Error_Exit( $ThisFunc, "Own the Outlook Singleton Mutex", Default, "Another process already owns the Outlook Singleton Mutex." )
	EndIf

EndFunc

;	Release my ownership of the Outlook Singleton Mutex
;
;		Will be done automatically when the script exits.
;
;		This is for when you want to let go of the Singleton but continue the script
;		so you can do things that don't require the Outlook app to perform.
;
Func Outlook_Singleton_Release( $hMutex )
Local Const $ThisFunc = "Outlook_Singleton_Release"

	If $hMutex = 0 Then Return
	If Not _WinAPI_CloseHandle( $hMutex ) Then hr_Error_Exit( $ThisFunc, "_WinAPI_CloseHandle", _WinAPI_GetLastError(), "Failed to release the Outlook Singleton Mutex." )

EndFunc
