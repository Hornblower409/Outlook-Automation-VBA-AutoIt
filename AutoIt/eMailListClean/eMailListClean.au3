#include "HotRod\hr_Directives.au3"
#include "HotRod\hr_Script_Error_Exit.au3"
#include "HotRod\hr_Debug_Print.au3"

#include <StringConstants.au3>

; =====================================================================
;	Clean up a list of email addresses on the Clipboard
; =====================================================================

	; Get the Clipboard text

		Local $List_Original = ClipGet( )
		If @error Then
			hr_Script_Error_Exit( "Get the Clipboard text", "Returned an @error", @error )
		EndIf
		Local $List = $List_Original
		; hr_Debug_Print( @ScriptName, "Get the Clipboard text", "$List_Original", $List_Original )

	; Replace any <CR> with Semicolon

		$List = StringReplace( $List, @CR, ";" )
		; hr_Debug_Print( @ScriptName, "Replace any <CR> with Semicolon", "$List", $List )

	; Strip all Whitespace

		$List = StringStripWS( $List, $STR_STRIPALL )

	; Replace commas with semicolons
	; Colapse consecutive semicolons

		$List = StringReplace( $List, ",", ";" )
		$List = StringReplace( $List, ";;", ";" )
		; hr_Debug_Print( @ScriptName, "All Replacement Done", "$List", $List )

	; Remove Leading or Trailing semicolons

		If StringLeft(  $List, 1 ) = ";" Then $List = StringTrimLeft(  $List, 1 )
		If StringRight( $List, 1 ) = ";" Then $List = StringTrimRight( $List, 1 )
		hr_Debug_Print( @ScriptName, "Clean Up Done", "$List", $List )

	; Error if the final list is emptyn

		If ($List = "")  Then
			hr_Script_Error_Exit( "Check Final List", "Final List is empty.", Default, "Original List", $List_Original )
		EndIf

	; RegExp (God Help Us) the list for valid addresses
	;
	;	From https://www.autoitscript.com/forum/topic/150225-validate-if-user-entering-the-right-email-address-format/
	;

		Local $localpart = "[[:alnum:]!#$%&'*+-/=?^_`{|}~.]+"
		Local $domainname = "[[:alnum:].-]+\.[[:alnum:]]+"

		Local $Adrs = StringSplit($List, ';')


		For $Inx = 1 To $Adrs[0]

			If NOT (StringRegExp($Adrs[$Inx], '(?i)^(' & $localpart & ')@(' & $domainname & ')$', 0)) Then
				hr_Script_Error_Exit( "RegExp email addresses", "Invlaid email address.", Default, "Address", $Adrs[$Inx] )
			EndIf

		Next ; $Inx

	; Put the clean list on the Clipboard

		ClipPut( $List )

	Exit( 0 )
