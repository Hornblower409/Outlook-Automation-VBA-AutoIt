#include-once

#include "HotRod\hr_Keyboard_ModifiersUp.au3"

; =====================================================================
; 2022-09-25
;
;	Send a Alt sequence to move around the Ribbion in the style that Stupid likes.
;
;	Calling:
;
;		Local $AltSeq = [ "{char}", "{char}", ... ]
;		Outlook_AltSeqSend( $AltSeq, Optional $Delay between Sends )
;
; =====================================================================
Func Outlook_AltSeqSend( $AltSeq, $Delay = 100 )

	;	Wait for all Modifier Keys Released
	;
	;		Or Stupid will not see my {ALTDOWN} {ALTUP}
	;
	hr_Keyboard_ModifiersUp( )

	Send( "{ALTDOWN}" )
	Send( "{ALTUP}" )

	For $i = 0 To UBound( $AltSeq ) - 1
		Send( $AltSeq[$i] )
		Sleep( $Delay )
	Next

EndFunc

