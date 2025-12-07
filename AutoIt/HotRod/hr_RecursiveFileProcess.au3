#include-once

; =====================================================================
; 2022-08-26
;
;	Walk a directory tree and process every file.
;
;	Include this file and then provide the process function as: hr_CallBack_RecursiveFileProcess( $Path )
;	See C:\HotRod\AU3\System\Link\Icons\Broken for an example
;
;	From  https://www.autoitscript.com/forum/topic/58558-filelisttoarray-comprehensive-comparison/
;
; =====================================================================
Func hr_RecursiveFileProcess($startDir, $depth = 0)
Local Const $ThisFunc = "hr_RecursiveFileProcess"

	Local $search, $next

    $search = FileFindFirstFile($startDir & "\*.*")
    If @error Then Return

	While 1
        $next = FileFindNextFile($search)
        If @error Then ExitLoop

        If StringInStr(FileGetAttrib($startDir & "\" & $next), "D") Then
            hr_RecursiveFileProcess($startDir & "\" & $next, $depth + 1)
        Else
            hr_CallBack_RecursiveFileProcess( $startDir & "\" & $next )
        EndIf
	WEnd

    FileClose($search)

EndFunc

; ---------------------------------------------------------------------
; Func  hr_CallBack_RecursiveFileProcess( $Path )
; Local Const $ThisFunc = "hr_CallBack_RecursiveFileProcess"
;
;
; EndFunc
