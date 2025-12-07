#include-once

#include "HotRod\hr_Error_Exit.au3"
#include "HotRod\hr_Misc.au3"
#include "HotRod\hr_ExpandEnvStrings.au3"

#include <File.au3>
#include <WinAPIFiles.au3>

; =====================================================================
;	Open a Windows Explorer on a path
; =====================================================================
Func hr_Explorer_OpenPath( $Path )
Local Const $ThisFunc = "hr_Explorer_OpenPath"

	;	If "Quoted" - remove them

	If ( StringLeft( $Path, 1 ) = $Quote ) And ( StringRight( $Path, 1 ) = $Quote ) Then
		$Path = StringMid( $Path, 2, StringLen( $Path ) - 2 )
	EndIf

	;	Do Environment Variable expansion on the Path
	;
	$Path = hr_ExpandEnvStrings( $Path )

	;	Trim Whitespace from Path
	;
	$Path = StringStripWS( $Path,  $STR_STRIPLEADING + $STR_STRIPTRAILING )

	;	Get the real path to Sysnative for later check
	;
	;		2017-07-24 - http://www.samlogic.net/articles/sysnative-folder-64-bit-windows.htm
	;		Sysnative is a virtual folder, a special alias, that can be used to access the 64-bit System32 folder from a 32-bit application or script.
	;
	Local $SysNative = hr_ExpandEnvStrings( "%SystemRoot%\sysnative" )

	;	Convert any Relative path to absolute based on C:\
	;
	;	2023-02-22 - Bypass for network paths (// or \\)
	;
	Local $PathFull
	If StringLeft($Path, 2) = "//" Or StringLeft($Path, 2) = "\\" Then
		$PathFull = $Path
	Else
		$PathFull = _PathFull( $Path, "C:\" )
	EndIf

	;	Check for Sysnative (Param starts with "%SystemRoot%\sysnative")
	;
	If StringUpper( StringLeft( $PathFull, StringLen( $SysNative ) ) ) = StringUpper( $SysNative ) Then
		hr_Error_Exit( $ThisFunc, "Check for Sysnative", Default, "Path is to a Windows Redirection folder", "$PathFull", $PathFull )
	EndIf

	;	Check that the path exist and get it's Attributes (Type)
	;
	;	2023-02-22 - Bypass for WSL: \\wsl.localhost and \\wsl$
	;
	Local $PathAttrb
	If $PathFull = "\\wsl.localhost" Or $PathFull = "\\wsl$" Then
		$PathAttrb = $FILE_ATTRIBUTE_DIRECTORY
	Else
		$PathAttrb = _WinAPI_GetFileAttributes( $PathFull )
		If $PathAttrb = 0 Then hr_Error_Exit( $ThisFunc, "GetFileAttributes", Default, "Path does not exist", "$Path", $Path, "$PathFull", $PathFull )
	EndIf

	; Build Explorer command to open the Path
	;
	Local $RunCmd

	;	If a Directory - Open that dir
	;	Else - Open the dir and select the file
	;
	If BitAND( $PathAttrb, $FILE_ATTRIBUTE_DIRECTORY ) Then
		$RunCmd = 'explorer /n,"' & $PathFull & '"'
	Else
		$RunCmd = 'explorer /n,/select,"' & $PathFull & '"'
	EndIf

	;	Run the Explorer command

	Local $RunResults = RunWait( $RunCmd )
	If @error Then hr_Error_Exit( $ThisFunc, "Run Explorer Command", @error, "Failed", "$RunCmd", $RunCmd )
	If NOT $RunResults  Then hr_Error_Exit( $ThisFunc, "Run Explorer Command", $RunResults, "Returned non zero", "$RunCmd", $RunCmd )

EndFunc

