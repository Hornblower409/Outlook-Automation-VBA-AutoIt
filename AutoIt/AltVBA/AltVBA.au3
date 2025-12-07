#include "HotRod\hr_Directives.au3"

#include "Outlook\AutomationInit.au3"
#include "Outlook\Globals.au3"
#include "Outlook\EXEs.au3"

;	Show AutoIt Tray Icon
;
; #NoTrayIcon

; =====================================================================
;	2023-07-26
;
;	Start Outlook using an alternative (non default) VBAProject.otm
;
; =====================================================================
Local Const $ThisFunc = "Main"


	;	Outlook Automation Init
	;
	;	Must get a Results = 3 - Outlook is not running
	;
	Local $OutlookAppObject, $OutlookSingletonMutex
	Local $Results = Outlook_AutomationInit( $OutlookAppObject, $OutlookSingletonMutex )
	If Number( $Results ) = 0 Then hr_Error_Exit( $ThisFunc, "Outlook_AutomationInit", Default, "Outlook is already running.")
	If Number( $Results ) <> 3 Then hr_Error_Exit( $ThisFunc, "Outlook_AutomationInit", Default, "Outlook Automation Init failed", "$Results", $Results)

	;	Get a OTM to start
	;
	Local $AltVBAFile = File_Select()

	;	Set an Environment Variable for Outlook VBA Application_Startup to see
	;	with the path of the OTM file
	;
	EnvSet( $Outlook_Global_AltVBAEnv , $AltVBAFile )

	;	Start Outlook with the OTM and wait for completion
	;
	;		Can not let any other AutoIt Outlook scripts run while we are using an AltVBA
	;
	Local $Param = "/altvba " & '"' & $AltVBAFile & '"'
	ShellExecuteWait( $OutlookCmd, $Param )

Exit( 0 )

; ---------------------------------------------------------------------
;	File Select
;
;		Derived from AutoIt FileOpenDialog Help Example
;
; ---------------------------------------------------------------------
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <String.au3>

Func File_Select()
Local Const $ThisFunc = "File_Select"

	Local Const $sTitle = "."

	; Do a open file dialog
	;
	Local $Title = "Select the VBAProject OTM to start."
	$Title = $Title & _StringRepeat(" ", 25) & "(Program: " & @ScriptFullPath & ")"
	Local $InitDir = $Outlook_Global_VBAProjectBackupFolder
	Local $TypeFilter = "Outlook VBAProject OTM (*.otm)"
	Local $Options = BitOR($FD_FILEMUSTEXIST, $FD_PATHMUSTEXIST)
	Local $sFileOpenDialog = FileOpenDialog($Title, $InitDir & "\", $TypeFilter, $Options)
	If @error Then hr_Error_Exit( $ThisFunc, "FileOpenDialog", Default, "No Outlook VBAProject OTM file selected.")

	;	Return the full path to the selected file
	;
	Return $sFileOpenDialog

EndFunc