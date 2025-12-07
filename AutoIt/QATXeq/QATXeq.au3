#include "HotRod\hr_Directives.au3"
#include "HotRod\hr_Error_Exit.au3"
#include "HotRod\hr_Misc.au3"
#include "HotRod\hr_Debug_Print.au3"

#include "Outlook\EXEs.au3"

#NoTrayIcon

; =====================================================================
; 2024-11-08 - Execute a QAT item by Number
;
;	QATNumber	<-	0 thru $QATMaxNumber
;
;		0 = Do initial HotRodQATXeq Window Setup.
;		N = Execute that QAT item.
;
;	This is for when I need to run an Outlook Macro but ProcXeq is
;	not available. Like when the VBE has killed all my Global Objects
;	and unhooked all my events.
;
;	This is low level. No parameter passing.
;
;	Makes the Desktop flicker as I bring up my invisible, hidden window so I
;	can get to a QAT.
;
;	Adding a new Macro
;
;		Add the macro to the next avaiable QAT position.
;		Update $QATMaxNumber below to the new number of QAT items.
;
; =====================================================================

	;	Max QAT Number defined in Outlook
	;
	Global Const $QATMaxNumber = 1

	;	Original Outlook Journal Window
	;
	Global Const $OutlookJournal_CmdLine = '/select "Outlook:Journal"'
	Global Const $OutlookJournal_AWD = "[TITLE:Journal - Default - Microsoft Outlook;CLASS:rctrl_renwnd32]"

	;	After I've changed the Journal Window to my HotRodQATXeq Window
	;
	Global Const $HotRodQATXeq_Title = "HotRodQATXeq"
	Global Const $HotRodQATXeq_AWD = "[TITLE:" & $HotRodQATXeq_Title & ";CLASS:rctrl_renwnd32]"
	Global $HotRodQATXeq_HWnd
	Global Const $HotRodQATXeqControl_AWD = "[CLASS:NetUIHWND; INSTANCE:1]"
	Global $HotRodQATXeqControl_HWnd

	;	Window that was Active when I started
	;
	Global $LastActive_HWnd
	Global $LastActive_Title

	;	QAT Number to Execute from the Command Line
	;
	Global $QATNumber

Main()
Func Main()

	GetLastActive()

	CmdLine()
	If $QATNumber = 0 Then
		Window_Setup()
	Else
		QAT_Send()
	EndIf

	RestoreLastActive()

Exit( 0 )
EndFunc

Func CmdLine()
Local Const $FuncName = "CmdLine"

	;	Must be one, and only one, command line arg
	;
	If $CmdLine[0] <> 1 Then
		hr_Error_Exit( $FuncName, "Command Line argument count check", Default, _
		"Command Line must have one, and only one, argument." )
	EndIf

	;	!SPOS! Passing "" counts as a param
	;
	If $CmdLine[1] = "" Then
		hr_Error_Exit( $FuncName, "Command Line null argument check", Default, _
		"The Command Line argument can not be an empty string." )
	EndIf

	;	Must be a Number between 0 and QATMaxNumber
	;
	$QATNumber = $CmdLine[1]

	If StringIsDigit( $QATNumber ) = 0 Then
		hr_Error_Exit( $FuncName, "Command Line argument number check", Default, _
		"The Command Line argument must be a whole number." )
	EndIf

	If ( $QATNumber < 0 ) Or ( $QATNumber > $QATMaxNumber ) Then
		hr_Error_Exit( $FuncName, "Command Line argument range check", Default, _
		"The Command Line argument must be between 0 and QATMaxNumber.", _
		"$QATMaxNumber", $QATMaxNumber)
	EndIf

EndFunc

;	Get the Window that was Active when I started
;
Func GetLastActive()

	$LastActive_HWnd = WinGetHandle( "[Active]" )
	If @error Then
		$LastActive_HWnd = 0
		$LastActive_Title = ""
	EndIf

EndFunc

Func RestoreLastActive()
Local Const $FuncName = "RestoreLastActive"

	;	If no Last Active - Done
	;
	If $LastActive_HWnd = 0 Then Return

	;	Make it Active
	;
	If WinActivate( $LastActive_HWnd ) = 0 Then
		hr_Error_Exit( $FuncName, DEFAULT, DEFAULT, _
			"Last Active Window not found or would not Activate.", _
			"$LastActive_Title", $LastActive_Title )
	EndIf

	If WinWaitActive( $LastActive_HWnd, "", 5 ) = 0 Then
		hr_Error_Exit( $FuncName, DEFAULT, DEFAULT, _
			"Last Active Window did not become Active.", _
			"$LastActive_Title", $LastActive_Title )
	EndIf

EndFunc

;	Execute a QAT by Number
;
Func QAT_Send(  )
Local Const $FuncName = "QAT_Send"

	;	Get the Handle of the existing HotRodQATXeq Window or Create a new one
	;
	Window_Setup()

	;	Show the Window (it's off screen)
	;
	If WinSetState( $HotRodQATXeq_HWnd, "", @SW_Show ) = 0 Then
		hr_Error_Exit( $FuncName, DEFAULT, DEFAULT, _
			"HotRodQATXeq Window would not Show.", _
			"$HotRodQATXeq_AWD", $HotRodQATXeq_AWD )
	EndIf

	;	Make it Active
	;
	If WinActivate( $HotRodQATXeq_HWnd ) = 0 Then
		hr_Error_Exit( $FuncName, DEFAULT, DEFAULT, _
			"HotRodQATXeq Window would not Activate.", _
			"$HotRodQATXeq_AWD", $HotRodQATXeq_AWD )
	EndIf

	If WinWaitActive( $HotRodQATXeq_HWnd, "", 5 ) = 0 Then
		hr_Error_Exit( $FuncName, DEFAULT, DEFAULT, _
			"HotRodQATXeq Window did not become Active.", _
			"$HotRodQATXeq_AWD", $HotRodQATXeq_AWD )
	EndIf

	;	Send it the keys to execute the QAT
	;
	If ControlSend($HotRodQATXeq_HWnd, "", $HotRodQATXeqControl_HWnd, "!" & $QATNumber ) = 0 Then
		hr_Error_Exit( $FuncName, DEFAULT, DEFAULT, _
			"HotRodQATXeq Window would not take a ControlSend.", _
			"$HotRodQATXeq_AWD", $HotRodQATXeq_AWD )
	EndIf

	;	Hide it
	;
	If WinSetState( $HotRodQATXeq_HWnd, "", @SW_HIDE ) = 0 Then
		hr_Error_Exit( $FuncName, DEFAULT, DEFAULT, _
			"HotRodQATXeq Window would not Hide.", _
			"$HotRodQATXeq_AWD", $HotRodQATXeq_AWD )
	EndIf

EndFunc

;	Setup a HotRodQATXeq Window if one does not exist.
;
Func Window_Setup()
Local Const $FuncName = "Window_Setup"

	;	If the HotRodQATXeq Window already exist - Get the Handles and Done
	;
	$HotRodQATXeq_HWnd = WinGetHandle( $HotRodQATXeq_AWD )
	If Not @error Then
		$HotRodQATXeqControl_HWnd = ControlGetHandle( $HotRodQATXeq_HWnd, "", $HotRodQATXeqControl_AWD )
		If @error Then
			hr_Error_Exit( $FuncName, DEFAULT, @error, _
			"Got HotRodQATXeq Window Handle but not the QAT Control Handle.", _
			"$HotRodQATXeq_AWD", $HotRodQATXeq_AWD, "$HotRodQATXeqControl_AWD", $HotRodQATXeqControl_AWD )
		Else
			Return
		EndIf
	EndIf

	;	Open a new Journal Explorer Window
	;
	ShellExecute( $OutlookCmd, $OutlookJournal_CmdLine )
	If @error Then
		hr_Error_Exit( $FuncName, DEFAULT, @error, _
		"Open new Journale Window ShellExecute Failed", _
		"$OutlookCmd", $OutlookCmd, "$OutlookJournal_CmdLine", $OutlookJournal_CmdLine )
	EndIf

	;	Wait for it to Exist
	;
	$HotRodQATXeq_HWnd = WinWait( $OutlookJournal_AWD, "", 5 )
	If $HotRodQATXeq_HWnd = 0 Then
		hr_Error_Exit( $FuncName, DEFAULT, DEFAULT, _
			"Outlook Journal Window failed to Open.", _
			"$OutlookJournal_AWD", $OutlookJournal_AWD )
	EndIf

	;	Change the Title to my HotRodQATXeq Window
	;
	If WinSetTitle( $OutlookJournal_AWD, "", $HotRodQATXeq_Title ) = 0 Then
		hr_Error_Exit( $FuncName, DEFAULT, DEFAULT, _
			"Failed to change Outlook Journal Window Title to HotRodQATXeq.", _
			"$OutlookJournal_AWD", $OutlookJournal_AWD )
	EndIf

	;	Make SURE the Title Changed
	;
	If WinExists($HotRodQATXeq_AWD ) = 0 Then
		hr_Error_Exit( $FuncName, DEFAULT, DEFAULT, _
			"HotRodQATXeq Window does not exist. (Journal title change failed?)", _
			"$HotRodQATXeq_AWD", $HotRodQATXeq_AWD, "$OutlookJournal_AWD", $OutlookJournal_AWD )
	EndIf

	;	Hide it
	;
	If WinSetState( $HotRodQATXeq_HWnd, "", @SW_HIDE ) = 0 Then
		hr_Error_Exit( $FuncName, DEFAULT, DEFAULT, _
			"HotRodQATXeq Window would not Hide.", _
			"$HotRodQATXeq_AWD", $HotRodQATXeq_AWD )
	EndIf

;~ 	;	Make it Invisible
;~ 	;
;~ 	If WinSetTrans( $HotRodQATXeq_HWnd, "", 0 ) = 0 Then
;~ 		hr_Error_Exit( $FuncName, DEFAULT, DEFAULT, _
;~ 			"HotRodQATXeq Window would not become Invisible.", _
;~ 			"$HotRodQATXeq_AWD", $HotRodQATXeq_AWD )
;~ 	EndIf

;~ 	;	Resize it to the minimum that Stupid will still allow
;~ 	;	me to execute a QAT. Any smaller and he ignores me.
;~ 	;
;~ 	If WinMove( $HotRodQATXeq_HWnd, "", 0, 0, 326, 260 ) = 0 Then
;~ 		hr_Error_Exit( $FuncName, DEFAULT, DEFAULT, _
;~ 			"HotRodQATXeq Window would Resize.", _
;~ 			"$HotRodQATXeq_AWD", $HotRodQATXeq_AWD )
;~ 	EndIf

	;	Move it off screen
	;
	If WinMove( $HotRodQATXeq_HWnd, "", @DesktopWidth, @DesktopHeight ) = 0 Then
		hr_Error_Exit( $FuncName, DEFAULT, DEFAULT, _
			"HotRodQATXeq Window would Resize.", _
			"$HotRodQATXeq_AWD", $HotRodQATXeq_AWD )
	EndIf


EndFunc
