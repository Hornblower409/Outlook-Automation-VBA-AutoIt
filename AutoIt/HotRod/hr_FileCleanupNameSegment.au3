#include-once

; =====================================================================
;	2024-11-24 - AutoIt version of my Outlook VBA File_CleanupNameSegment
;
;   Replace any Invalid Characters in a Piece of a FileSpec
;
;       NOT for a full File Spec. Only pieces after "C:" and between "/\"s.
;       See: https://learn.microsoft.com/en-us/windows/win32/fileio/naming-a-file
;       2024-10-29 - Added any Control Chars
;
Func hr_FileCleanupNameSegment( ByRef Const $Raw, $RepChar = "_")

    Local $Cooked
    $Cooked = $Raw

    ;   Replace any Control Chars
    ;
    Local $AscV
    Local $LoopIx
    For $LoopIx = 1 To StringLen($Cooked)
        $AscV = Asc(StringMid($Cooked, $LoopIx, 1))
        Switch $AscV
            Case 0 To 31, 127, 251 To 255
                $Cooked = StringReplace($Cooked, Chr($AscV), $RepChar)
            Case Else
            ; Continue
        EndSwitch

    Next ;  LoopIx

    ;   Replace any Invalids
    ;
    Local $Invalids
    $Invalids = "<>:""/\|?*"

    For $LoopIx = 1 To StringLen($Invalids)
        $Cooked = StringReplace($Cooked, StringMid($Invalids, $LoopIx, 1), $RepChar)
    Next ;  LoopIx

    Return $Cooked

EndFunc
