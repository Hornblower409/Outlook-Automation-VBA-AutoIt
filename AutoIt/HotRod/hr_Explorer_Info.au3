#include-once

#include "HotRod\hr_Error_Exit.au3"
#include "HotRod\hr_COM_Error_Exit.au3"

#include <Array.au3>

; ---------------------------------------------------------------------
;	Test Jig
; ---------------------------------------------------------------------
Func hr_Explorer_Info_TEST()
Local Const $ThisFunc = "hr_Explorer_Info_TEST"

	Local $EWins = hr_Explorer_Info( 3 )

	For $i = 1 To $EWins[0]

		_ArrayDisplay( ($EWins[$i]) )

		If ($EWins[$i])[0] = 3 Then
			_ArrayDisplay( ($EWins[$i])[3] )
		EndIf

	Next

EndFunc

; =====================================================================
;	Returns an array of info on all open Explorer windows.
;
;	Option: 1 = Path Only, 2 = Path + Selected, 3 = Path + All
;
;	$E[0]			= Number of Explorer windows
;	$E[1+]			= $W[]
;		$W[0]		= 2 or 3
;		$W[1]		= Path
;		$W[2]		= Handle
;		$W[3]		= $L[]	Only If Option > 1
;			$L[0]	= Number of Line Items
;			$L[1+]	= Line Item
;
;---------------------------------------------------------------------
;	Derived From:
;	https://www.autoitscript.com/forum/topic/89833-windows-explorer-current-folder/?do=findComment&comment=973904
;	March 25, 2012	Melba23
;=====================================================================
Func hr_Explorer_Info( $Option )
Local Const $ThisFunc = "hr_Explorer_AllOpen"

	;	Get an array of all open Explorer Window Titles and Handles
	;
	Local $WinList = WinList("[REGEXPCLASS:^(Explore|Cabinet)WClass$]")
	Local $WinCount = $WinList[0][0]

	;	Init the Return Array
	;
	Dim $Return[$WinCount + 1]
	$Return[0] = $WinCount

	;	Walk the WinList
	;
	For $i = 1 To $WinCount

		;	Init the Window array
		;
		Dim $Win[3]
		$Win[0] = 2

		;	Get an arrary of the window's path and optionally selected/all line items
		;
		Local $WinInfo = hr_Explorer_Info_WinGet($WinList[$i][1], $Option)

		;	Put the Path and Handle in the Window array
		;
		$Win[1] = $WinInfo[1]
		$Win[2] = $WinList[$i][1]

		;	If any Line Items - put them in a Lines array
		;
		Local $LineCount = $WinInfo[0] - 1
		If $LineCount > 0 Then

			Dim $Lines[$LineCount + 1]
			$Lines[0] = $LineCount
			For $j = 2 To $WinInfo[0]
				$Lines[$j -1] = $WinInfo[$j]
			Next

			;	Put the Lines array into the Windows array
			;
			ReDim $Win[4]
			$Win[0] = 3
			$Win[3] = $Lines

		EndIf

		;	Put the Window array into the Return array
		;
		$Return[$i] = $Win

	Next

	Return SetExtended(0,$Return)

EndFunc

; ---------------------------------------------------------------------
; Func hr_Explorer_Info_WinGet($hWnd)
; Author: klaus.s, KaFu, Ascend4nt (consolidation & cleanup, Path name simplification)
; ---------------------------------------------------------------------
Func hr_Explorer_Info_WinGet($hWnd, $Option)
Local Const $ThisFunc = "hr_Explorer_Info_WinGet"
Local $hr_COM_Error_OnError[1] = [$ThisFunc]

    If Not IsHWnd($hWnd) Then Return hr_Error_Exit( $ThisFunc, "Check hWnd", Default, "Value is not a valid window handle.", "$hWnd", $hWnd )

	;	Get the FolderView object
	;
    Local $oSHFolderView = hr_Explorer_Info_FolderView($hWnd)
    If @error Then Return SetError(@error,0,'')

	;	Define the return array and get the Path
	;
	Local $Return[2] = [0, ""]
	$Return[0] = 1
	$Return[1] = $oSHFolderView.Folder.Self.Path

	;	If Only Path - we're done
	;
	If $Option = 1 Then  Return $Return

	;	$oLines = FolderView Only Selected (.SelectedItems) or All (.Folder.Items)
	;
	Dim $oLines
	If $Option = 2 Then $oLines = $oSHFolderView.SelectedItems
	If $Option = 3 Then $oLines = $oSHFolderView.Folder.Items

	;	ReDim Return for the number of Lines & adjust the Count
	;
	ReDim $Return[2 + $oLines.Count]
	$Return[0] = $oLines.Count + 1

	;	Copy the Lines into the Return array after the Count and Path
	;
	Local $iCounter = 2
	For $oFolderItem In $oLines
		$Return[$iCounter] = $oFolderItem.Path
		$iCounter += 1
	Next

	Return $Return

EndFunc

; ---------------------------------------------------------------------
; Func hr_Explorer_Info_FolderView($hWnd)
; Returns an 'ShellFolderView' Object for the given Window handle
; Author: Ascend4nt, based on code by KaFu, klaus.s
; ---------------------------------------------------------------------
Func hr_Explorer_Info_FolderView($hWnd)
Local Const $ThisFunc = "hr_Explorer_Info_FolderView"
Local $hr_COM_Error_OnError[1] = [$ThisFunc]

    If Not IsHWnd($hWnd) Then Return hr_Error_Exit( $ThisFunc, "Check hWnd", Default, "Value is not a valid window handle.", "$hWnd", $hWnd )
    Local $oShell,$oShellWindows,$oIEObject,$oSHFolderView

    ; Create a Shell Object
	;
    $oShell=ObjCreate("Shell.Application")
    If Not IsObj($oShell) Then hr_Error_Exit( $ThisFunc, "Create Shell.Application", Default, "Failed" )

	; Get a ShellWindows Collection object
	;
    $oShellWindows = $oShell.Windows()
    If Not IsObj($oShellWindows) Then hr_Error_Exit( $ThisFunc, "Get Shell.Windows", Default, "Failed" )

	; Iterate through the collection - each of type 'InternetExplorer' Object
	;
    For $oIEObject In $oShellWindows

        ; InternetExplorer->Document = ShellFolderView object
		;
		If $oIEObject.HWND = $hWnd Then
            $oSHFolderView=$oIEObject.Document
            If IsObj($oSHFolderView) Then Return $oSHFolderView
            hr_Error_Exit( $ThisFunc, "Get ShellFolderView object", Default, "Failed" )
        EndIf

    Next

    Return

EndFunc
