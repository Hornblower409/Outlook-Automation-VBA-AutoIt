#include-once

#include <Array.au3>

#include "HotRod\hr_Error_Exit.au3"

; =====================================================================
;	2023-07-31 - Registry Operations that mimic Wscript.Shell
;	so it can use the same RegKey strings as VBA calling Wscript.Shell
;
;	See:
;
;		RegRead Method		https://www.vbsedit.com/html/1b567504-59f4-40a9-b586-0be49ab3a015.asp
;		RegDelete Method	https://www.vbsedit.com/html/161db13b-c4ca-4aec-8899-697a4183e82c.asp
;		RegWrite Method		https://www.vbsedit.com/html/678e6992-ddc4-4333-a78c-6415c9ebcc77.asp
;
;		https://www.autoitscript.com/autoit3/docs/functions/RegRead.htm
;		https://www.autoitscript.com/autoit3/docs/libfunctions/WinAPIEx%20Registry%20Management.htm
;
; =====================================================================

	Func hr_Registry_Read( $RegEntry ) ; As String
	Local Const $ThisFunc = "hr_Registry_Read"

		Local $RegKeyName, $RegValueName, $RegValue
		If Not hr_Registry_SplitEntry( $RegEntry, $RegKeyName, $RegValueName ) Then hr_Error_Exit( $ThisFunc, "hr_Registry_SplitEntry", Default, "Call Failed." )

		$RegValue = RegRead( $RegKeyName, $RegValueName )
		If @error Then hr_Error_Exit( $ThisFunc, "RegRead", @error, "Reg Read failed.", "$RegKeyName", $RegKeyName, "$RegValueName", $RegValueName )
		Return $RegValue

	EndFunc

	Func hr_Registry_Delete( $RegEntry )
	Local Const $ThisFunc = "hr_Registry_Delete"

		Local $RegKeyName, $RegValueName
		If Not hr_Registry_SplitEntry( $RegEntry, $RegKeyName, $RegValueName ) Then hr_Error_Exit( $ThisFunc, "hr_Registry_SplitEntry", Default, "Call Failed." )

		;	Pick a method to do the Delete - Key or Value
		;
		Local $RegDelete
		If $RegValueName = "" Then
			$RegDelete = RegDelete($RegKeyName)
		Else
			$RegDelete = RegDelete($RegKeyName, $RegValueName)
		EndIf

		;	If results = key/value does not exist Or Success - done
		;
		If ($RegDelete = 0) Or ($RegDelete = 1) Then Return
		If @error Then hr_Error_Exit( $ThisFunc, "RegRead", @error, "Reg Delete failed.", "$RegKeyName", $RegKeyName, "$RegValueName", $RegValueName )

	EndFunc

	Func hr_Registry_Write( $RegEntry,  $RegValue = "", $RegType = "REG_SZ")
	Local Const $ThisFunc = "hr_Registry_Write"

		Local $RegKeyName, $RegValueName
		If Not hr_Registry_SplitEntry( $RegEntry, $RegKeyName, $RegValueName ) Then hr_Error_Exit( $ThisFunc, "hr_Registry_SplitEntry", Default, "Call Failed." )

		;	Pick a method to do the Write - Key or Value
		;
		If $RegValueName = "" Then
			RegWrite($RegKeyName)
		Else
			RegWrite($RegKeyName, $RegValueName, $RegType, $RegValue )
		EndIf
		If @error Then hr_Error_Exit( $ThisFunc, "RegWrite", @error, "Reg Write failed.", "$RegKeyName", $RegKeyName, "$RegValueName", $RegValueName, "$RegType", $RegType )

	EndFunc

	Func hr_Registry_SplitEntry( $RegEntry, ByRef $RegKeyName, ByRef $RegValueName ) ; As Boolean
	; Local Const $ThisFunc = "hr_Registry_SplitEntry"

		;	If the Entry name ends In a "\" (a key) then
		;		Value Name = "" (DEFAULT)
		;		Strip the trailing "\" off the Entry
		;
		If StringRight($RegEntry, 1) = "\" Then

			$RegKeyName = StringLeft( $RegEntry, StringLen( $RegEntry) - 1 )
			$RegValueName = ""

		;	Else
		;		Value Name = last part of the Entry Name
		;		Strip the Value Name off the Entry
		;
		Else

			Local $RegEntrySplit = StringSplit( $RegEntry, "\", $STR_ENTIRESPLIT )
			$RegValueName = $RegEntrySplit[$RegEntrySplit[0]]
			ReDim $RegEntrySplit[UBound($RegEntrySplit) - 1]
			$RegKeyName = _ArrayToString($RegEntrySplit, "\", 1 )

		EndIf

		Return True

	EndFunc