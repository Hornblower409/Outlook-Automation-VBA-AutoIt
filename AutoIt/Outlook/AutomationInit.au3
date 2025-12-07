#include-once

#include "HotRod\hr_Error_Exit.au3"
#include "HotRod\hr_Misc.au3"
#include "HotRod\hr_RegReadWrite.au3"
#include "HotRod\hr_System_StartupReg.au3"
#include "HotRod\hr_COM_Error_exit.au3"

#include "Outlook\Singleton.au3"
#include "Outlook\Globals.au3"

; ---------------------------------------------------------------------
;	2023-08-15
;
;	Standard preamble code for any Outlook automation
;
;	Numbers at the start of the error messages are used by the caller
;	to determine the error type. You can change the error message text
;	but don't change the numbers.
;
;	Usage:
;
;		#include "Outlook\AutomationInit.au3"
;
;		;	Outlook Automation preamble - exits if not successful
;		;
;		Local $OutlookAppObject, $OutlookSingletonMutex
;		Outlook_AutomationInit_Exit( $OutlookAppObject, $OutlookSingletonMutex )
;
;	When run Elevated will return a 3 (Outlook not running) because of UIPI.
;
; ---------------------------------------------------------------------
Func Outlook_AutomationInit( ByRef $OutlookAppObject, ByRef $OutlookSingletonMutex )
;~ Local Const $ThisFunc = "Outlook_AutomationInit"

	;	If System_Startup is running
	;
	Local $Startup_Status
	$Startup_Status = hr_RegRead( $System_Startup_RegKeyBase, $System_Startup_RegValueStatus )
	If $Startup_Status <> $System_Startup_RegValueStatusDone Then Return "1 - HotRod System_Startup is running"

	;	Get ownership of the Outlook Singleton Mutex
	;
	$OutlookSingletonMutex = Outlook_Singleton_Wait()
	If $OutlookSingletonMutex = 0 Then Return "2 - Timeout waiting for ownership of the Outlook Singleton Mutex"

	;	Get the Outlook App Object
	;
	Local $Caller_hr_COM_Error_Exit = $hr_COM_Error_Exit
	$hr_COM_Error_Exit = False
	$OutlookAppObject = ObjGet("", "Outlook.Application")
	$hr_COM_Error_Exit = $Caller_hr_COM_Error_Exit
	If Not IsObj($OutlookAppObject) Then Return "3 - Outlook application is not running"

	Return "OK" ; Number("OK") = 0

EndFunc

;	Silent Exit this script if AutomationInit fails
;
Func Outlook_AutomationInit_Exit( ByRef $OutlookAppObject, ByRef $OutlookSingletonMutex )
;~ Local Const $ThisFunc = "Outlook_AutomationInit_Exit"

	Local $Results = Outlook_AutomationInit( $OutlookAppObject, $OutlookSingletonMutex )
	If Number( $Results ) <> 0 Then Exit( 1 )

EndFunc

;	Show error message and exit this script if AutomationInit fails
;
Func Outlook_AutomationInit_ErrorExit( ByRef $OutlookAppObject, ByRef $OutlookSingletonMutex )
Local Const $ThisFunc = "Outlook_AutomationInit_ErrorExit"

	Local $Results = Outlook_AutomationInit( $OutlookAppObject, $OutlookSingletonMutex )
	If Number( $Results ) <> 0  Then hr_Error_Exit( $ThisFunc, "Outlook_AutomationInit", Default, "Outlook Automation Init failed", "$Results", $Results)

EndFunc