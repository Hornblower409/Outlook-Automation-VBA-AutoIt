#include-once

#include "HotRod\hr_Registry.au3"

#include "Outlook\AutomationInit.au3"
#include "Outlook\Globals.au3"
#include "Outlook\WinIsOutlook.au3"

; =====================================================================
; 2024-12-22
;
;	Function to Execute an Outlook VBA Proc with Param via Event
;
;	/NamedOptions MUST come before any arguments
;
;	$CommandLine can be an array like $CmdLine. e.g.:
;
;		Local $PXCommandLine[2] = [1, "ProcName"]
;		Local $PXCommandLine[4] = [3, "/Option, "ProcName", "Arg1"]
;		Local $PXCommandLine[3] = [2, "ProcName", "Arg1"]
;		Outlook_ProcXeq( $PXCommandLine )
;
;	Or just a Proc Name if there are no args or Options. e.g.:
;
;		Outlook_ProcXeq( "ConnectToServer" )
;
; =====================================================================
;
Func Outlook_ProcXeq( _
	$CommandLine, _
	$Option_OutlookWindowActive = Default, _
	$Option_SilentErrorExit = Default, _
	$Option_WaitReadyTimeout = Default, _
	$Option_WaitRunningTimeout = Default _
	) ; As Boolean

Local Const $ThisFunc = "Outlook_ProcXeq"

	;	Convert a simple $CommandLine string into an Array
	;
	If Not IsArray( $CommandLine ) Then
		Local $TempArray[2] = [1, $CommandLine]
		$CommandLine = $TempArray
	EndIf

	;	$CommandLine Array Check
	;
	If $CommandLine[0] < 1 Then hr_Error_Exit( $ThisFunc, "$CommandLine Check", Default, "$CommandLine Array can not be empty." )
	If $CommandLine[1] = "" Then hr_Error_Exit( $ThisFunc, "$CommandLine Check", Default, "$CommandLine[1] can not be an empty string" )

	;	Process any Named Options in the $CommandLine Array
	;
	While UBound( $CommandLine ) >  1

		;	If not a /Option - Done
		;
		If StringLeft( $CommandLine[1], 1 ) <> "/" Then ExitLoop

		;	For each /Option
		;
		Switch StringUpper( $CommandLine[1] )

			Case StringUpper("/NoOutlookWindowActive")
				$Option_OutlookWindowActive = False
				_ArrayDelete( $CommandLine, 1 )

			Case StringUpper("/SilentErrorExit")
				$Option_SilentErrorExit = True
				_ArrayDelete( $CommandLine, 1 )

			Case StringUpper("/WaitReadyTimeout")
				$Option_WaitReadyTimeout = $CommandLine[3]
				_ArrayDelete( $CommandLine, "1-2" )

			Case StringUpper("/WaitRunningTimeout")
				$Option_WaitRunningTimeout = $CommandLine[3]
				_ArrayDelete( $CommandLine, "1-2" )

			Case Else
				hr_Error_Exit( $ThisFunc, "Options Check", Default, "Invalid Option.", "Option String", $CommandLine[1] )

		EndSwitch

		If @error <> 0 Then hr_Error_Exit( $ThisFunc, "_ArrayDelete", @error  )
		$CommandLine[0] = UBound( $CommandLine ) - 1

	WEnd

	;	Must be at least one Command Line Arg left after Options removed
	;
	If $CommandLine[0] < 1 Then hr_Error_Exit( $ThisFunc, "After Options Removed Check", Default, "Command Line/Array must have at least one non-option arg." )

	;	Handle Default $Options
	;
	If $Option_OutlookWindowActive = Default Then $Option_OutlookWindowActive = True
	If $Option_SilentErrorExit = Default Then $Option_SilentErrorExit = False
	If $Option_WaitReadyTimeout = Default Then $Option_WaitReadyTimeout = 1000
	If $Option_WaitRunningTimeout = Default Then $Option_WaitRunningTimeout = 1000

	;	If required and the Active Window is not an Outlook Window - exit
	;
	If $Option_OutlookWindowActive Then

		Local $hWnd = WinGetHandle( "[Active]" )
		If @error Then hr_Error_Exit( $ThisFunc, "WinGetHandle", @error, "Failed", "Param", "[Active]" )
		If Not Outlook_WinIsOutlook( $hWnd ) Then hr_Error_Exit( $ThisFunc, "Outlook_WinIsOutlook", Default, "Active window is not an Outlook window." )

	EndIf

	;	Outlook Automation preamble - exit on failure
	;
	Local $OutlookAppObject, $OutlookSingletonMutex

	If $Option_SilentErrorExit Then
		Outlook_AutomationInit_Exit($OutlookAppObject, $OutlookSingletonMutex )
	Else
		Outlook_AutomationInit_ErrorExit( $OutlookAppObject, $OutlookSingletonMutex )
	EndIf

	;	If the last ProcXeq Canceled but Stupid never cleared it - Reset to Ready
	;
	Local $CurrentStatus = hr_Registry_Read( $Outlook_gblProcXeq_RegStatus )
	If $CurrentStatus = $Outlook_gblProcXeq_Status_Canceled Then hr_Registry_Write( $Outlook_gblProcXeq_RegStatus, $Outlook_gblProcXeq_Status_Ready )

	;	Wait for ProcXeq Status Ready - exit on failure
	;
	If Not Outlook_ProcXeq_StatusWait($CurrentStatus, $Outlook_gblProcXeq_Status_Ready, $Option_WaitReadyTimeout) Then
		If $Option_SilentErrorExit Then
			Exit( 1 )
		Else
			hr_Error_Exit( $ThisFunc, "Ready - Status Check", Default, "ProcXeq Status is not '" & $Outlook_gblProcXeq_Status_Ready & "'." & @CRLF & @CRLF & "Wait for any running ProcXeq to complete." & @CRLF & @CRLF & "- Try running VBA App Init -", "$CurrentStatus", $CurrentStatus )
		EndIf
	EndIf

	;	Get the Journal Item to be used as a trigger
	;	from the Entry/Store IDs stashed by VBA
	;
	Local $JEntryID = hr_Registry_Read( $Outlook_gblProcXeq_RegEntryId )
	Local $JStoreID = hr_Registry_Read( $Outlook_gblProcXeq_RegStoreId )
	Global $hr_COM_Error_OnError[2] = [$ThisFunc, "GetItemFromID($JEntryID, $JStoreID)"]
	Local $JItem = $OutlookAppObject.Session.GetItemFromID($JEntryID, $JStoreID)

	;	Write the ProcXeq Command Line to the Registry
	;
	_ArrayDelete( $CommandLine, 0 )
	Local $CommandString = _ArrayToString( $CommandLine, $Outlook_gblProcXeq_CmdLineSep )
	hr_Registry_Write( $Outlook_gblProcXeq_RegCmdLine, $CommandString )

	;	Set the Status to Submitted
	;
	hr_Registry_Write( $Outlook_gblProcXeq_RegStatus, $Outlook_gblProcXeq_Status_Submitted )

	;	Trigger the event
	;	Ignore any "Operation Canceled" or "Exception" error
	;
	Global $hr_COM_Error_OnError[2] = [$ThisFunc, "$JItem.Actions.Item($Outlook_gblProcXeq_ActionName).Execute"]
	Global $hr_COM_Error_Ignore[2] = ["0x80004004", "0x80010105"]
	$JItem.Actions.Item($Outlook_gblProcXeq_ActionName).Execute

	;	Wait for ProcXeq Status Running - cancel and exit on failure
	;
	If Not Outlook_ProcXeq_StatusWait($CurrentStatus, $Outlook_gblProcXeq_Status_Running, $Option_WaitRunningTimeout) Then

		hr_Registry_Write( $Outlook_gblProcXeq_RegStatus, $Outlook_gblProcXeq_Status_Canceled )

		If $Option_SilentErrorExit Then
			Exit( 1 )
		Else
			hr_Error_Exit( $ThisFunc, "Running - Status Check", Default, "ProcXeq Status did not change to '" & $Outlook_gblProcXeq_Status_Running & "'." & @CRLF & @CRLF & "- Try running VBA App Init -" & @CRLF & @CRLF & "Operation Canceled.", "$CurrentStatus", $CurrentStatus )
		EndIf
	EndIf

	Return True

EndFunc

;	Wait for a specific ProcXeq Status
;
Func Outlook_ProcXeq_StatusWait(ByRef $CurrentStatus, $RequiredStatus, $TimeOutMSecs ) ; As Boolean
;~ Local Const $ThisFunc = "Outlook_ProcXeq_StatusWait"

	Local $WaitedMSecs = 0
	Local Const $WaitCheckMSecs = 100

	Do
		$CurrentStatus = hr_Registry_Read( $Outlook_gblProcXeq_RegStatus )
		If  $CurrentStatus = $RequiredStatus Then Return True
		Sleep( $WaitCheckMSecs )
		$WaitedMSecs = $WaitedMSecs + $WaitCheckMSecs

	Until $WaitedMSecs > $TimeOutMSecs

	Return False

EndFunc

