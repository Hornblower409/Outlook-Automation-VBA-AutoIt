#include-once

#include "HotRod\hr_Error_Exit.au3"
#include "HotRod\hr_ClipGetText.au3"

#pragma compile(inputboxres, true)

; =====================================================================
; 2022-10-04
;
;	Check the Command Line for options:
;
;		/InputBox
;		/Clipboard
;		/Selection
;
;	If present:
;
;		The option must be the first and only Command Line arg.
;		Returns the results of performing that option.
;
;	Else the original $CmdLine[1]
;
;	Note to future self - Keep options simple and looking like normal Batch so that we
;	can use a real command line options parser in the future.
;
; =====================================================================
Func hr_CmdLine_Opt_BoxClipSel( ByRef Const $CmdLine, $Prompt = Default )
Local Const $ThisFunc = "hr_CmdLine_Opt_BoxClipSel"

	;	Must be one, and only one, command line arg
	;
	If $CmdLine[0] < 1 Then hr_Error_Exit( $ThisFunc, "Command Line argument count check", Default, "Command Line must have at least one argument" )

	;	!SPOS! Passing "" counts as a param
	;
	If $CmdLine[1] = "" Then hr_Error_Exit( $ThisFunc, "Command Line null argument check", Default, "The Command Line argument can not be an empty string" )

	;	If the InputBox Prompt is Default - use a generic
	;
	If $Prompt = Default Then $Prompt = ""

	;	Switch on the argument
	;
	Local $Rtn = ""

	Switch StringUpper( $CmdLine[1] )

		Case "/INPUTBOX"
			$Rtn = InputBox(@ScriptFullPath, $Prompt, Default, Default, 450, 130 )
			If $Rtn = "" Then hr_Error_Exit( $ThisFunc, "InputBox", @error, "No Input or Canceled"  )

		;	2025-07-07 - Started getting "Clipboard is not accessable"
		;	Added Caller and changed Delay from Default to 100
		;
		Case "/CLIPBOARD"
			$Rtn = hr_ClipGetText($ThisFunc, 100)

		Case "/SELECTION"
			Send( "^c" )
			$Rtn = hr_ClipGetText()

		Case Else
			$Rtn = $CmdLine[1]

	EndSwitch

	Return $Rtn

EndFunc
