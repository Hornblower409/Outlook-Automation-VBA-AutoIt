#include-once

#include "Outlook\TypeOfActiveWindow.au3"
#include "HotRod\hr_COM_Error_exit.au3"

; ---------------------------------------------------------------------
;	2024-11-18 - Active Window must be an Inspector or an Explorer with one item selected
;
;	AutoIt Version of my VBA Selected_SingleItemOnly
; ---------------------------------------------------------------------

Func Outlook_SingleItemOnly($OutlookAppObject, $OutlookSingletonMutex, ByRef $Item )
Local Const $ThisFunc = "Outlook_SingleItemOnly"

	Local $Caller_hr_COM_Error_Exit = $hr_COM_Error_Exit
	$hr_COM_Error_Exit = True
	With $OutlookAppObject

		Local $TypeOfActiveWindow = Outlook_TypeOfActiveWindow($OutlookAppObject, $OutlookSingletonMutex)

		If $TypeOfActiveWindow = "INSPECTOR" Then
			$Item = .ActiveInspector.CurrentItem
		ElseIf ($TypeOfActiveWindow = "EXPLORER") And ($OutlookAppObject.ActiveExplorer.Selection.Count = 1) Then
			$Item = .ActiveExplorer.Selection.Item(1)
		Else
			hr_Error_Exit( $ThisFunc, Default, Default, "The Outlook Active Window must be an Inspector or an Explorer with a single item selected." )
		EndIf

	EndWith
	$hr_COM_Error_Exit = $Caller_hr_COM_Error_Exit

EndFunc