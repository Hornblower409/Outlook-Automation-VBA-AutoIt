#include-once

#include "HotRod\hr_COM_Error_exit.au3"
#include "Outlook\Singleton.au3"

; =====================================================================
; 2023-09-05
;
;	From:
;
;		https://www.autoitscript.com/autoit3/docs/functions/ObjName.htm
;		https://www.autohotkey.com/boards/viewtopic.php?t=20407
;
;	Usage:
;
;		#include "Outlook\TypeOfActiveWindow.au3"
;
;		If you already have an Outlook App Object and SingletonMutex:
;
;			Outlook_TypeOfActiveWindow($OutlookAppObject, $OutlookSingletonMutex)
;
;		Else:
;
;			Outlook_TypeOfActiveWindow(Default, Default)
;
;
;	Return: "UNKNOWN", "EXPLORER", "INSPECTOR"
;
; =====================================================================

;~ ;	Unit Test
;~ ;
;~	Local Const $ThisFunc = "Main"
;~ 	Local $TypeOfActiveWindow = Outlook_TypeOfActiveWindow( Default, Default )
;~ 	hr_Debug_Print( $ThisFunc, Default, Default, "$TypeOfActiveWindow", $TypeOfActiveWindow )
;~ 	Exit ( 0 )

Func Outlook_TypeOfActiveWindow($OutlookAppObject = Default, $OutlookSingletonMutex = Default)

	Local $TypeOfActiveWindow = "UNKNOWN"
	Local $Caller_hr_COM_Error_Exit

	;	If no Outlook SingletonMutex from caller - Get ownership of it
	;
	Local $ReleaseOnReturn_SingletonMutex = False
	If $OutlookSingletonMutex = Default Then

		$OutlookSingletonMutex = Outlook_Singleton_Wait()
		If $OutlookSingletonMutex <> 0 Then $ReleaseOnReturn_SingletonMutex = True

	EndIf

	;	If no Outlook AppObject from caller - Get one
	;
	Local $ReleaseOnReturn_AppObject = False
	If $OutlookAppObject = Default Then

		$Caller_hr_COM_Error_Exit = $hr_COM_Error_Exit
		$hr_COM_Error_Exit = True

			$OutlookAppObject = ObjGet("", "Outlook.Application")

		$hr_COM_Error_Exit = $Caller_hr_COM_Error_Exit
		If IsObj($OutlookAppObject) Then $ReleaseOnReturn_AppObject = True

	EndIf

;	If I own the Mutex and have an App Object - Get the info
;
	If $OutlookSingletonMutex <> 0 And IsObj($OutlookAppObject) Then

		$Caller_hr_COM_Error_Exit = $hr_COM_Error_Exit
		$hr_COM_Error_Exit = True

			Local $ActiveWindow = $OutlookAppObject.ActiveWindow
			Local $ObjName = ObjName($ActiveWindow)
			Switch $ObjName
				Case "_Explorer"
					$TypeOfActiveWindow = "EXPLORER"
				Case "_Inspector"
					$TypeOfActiveWindow = "INSPECTOR"
				Case Else
					; Continue
			EndSwitch

		$hr_COM_Error_Exit = $Caller_hr_COM_Error_Exit

	EndIf

;	If not from the caller - Release the Outlook Object and the Singleton
;
	If $ReleaseOnReturn_AppObject Then $OutlookAppObject = 0
	If $ReleaseOnReturn_SingletonMutex Then Outlook_Singleton_Release( $OutlookSingletonMutex )

	Return $TypeOfActiveWindow

EndFunc