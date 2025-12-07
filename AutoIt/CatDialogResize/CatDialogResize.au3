#include "HotRod\hr_Directives.au3"
#include "HotRod\hr_Error_Exit.au3"

#NoTrayIcon

#include <GuiListView.au3>

; =====================================================================
;	Resize the Outlook Categories Dialog Name Column
; =====================================================================

	; Cat Dialog Window

		Global Const $CatWin_AWD = "[TITLE:Color Categories;CLASS:#32770]"
		Global $hCatWin

	; ListView Control (Cat List)

		Global Const $ListViewID = "[ID:4640]"

	; =====================================================================
	; MAIN
	; =====================================================================

		Wait_CatWindow( )
		Resize_ListView( )
		WinActivate ( $hCatWin )
		; Wait_CatWindowClose( )

	Exit( 0 )

; ---------------------------------------------------------------------
Func Resize_ListView( )
Local $FuncName = "Resize_ListView"

	; Get the ListView Control

	Local $hListView
	$hListView = ControlGetHandle( $hCatWin, "", $ListViewID )
	If $hListView = 0  Then
		hr_Error_Exit( $FuncName, DEFAULT, DEFAULT, _
			"Outlook Color Categories Dialog window - ListView Control not found.", _
			"$hCatWin", $hCatWin, "$ListViewID", $ListViewID )
	EndIf

	; Resize the ListView First Column (Name)

	If ( _GUICtrlListView_SetColumnWidth( $hListView, 0, 500 ) = 0 ) Then
		hr_Error_Exit( $FuncName, DEFAULT, DEFAULT, _
			'Outlook Color Categories Dialog window - ListView Control Col would not resize.', _
			"$hCatWin", $hCatWin, "$ListViewID", $ListViewID )
	EndIf


EndFunc

; ---------------------------------------------------------------------
Func Wait_CatWindow( )
Local $FuncName = "Wait_CatWindow"

	; Wait for the Cat Window to open

	$hCatWin = WinWait( $CatWin_AWD, "", 3 )
	If $hCatWin = 0  Then
		hr_Error_Exit( $FuncName, DEFAULT, DEFAULT, _
			"Outlook Color Categories Dialog window did not open.", _
			"$CatWin_AWD", $CatWin_AWD )
	EndIf

EndFunc

Func Wait_CatWindowClose( )
Local $FuncName = "Wait_CatWindowClose"

	WinWaitClose ( $hCatWin )

EndFunc