#include-once

; =====================================================================
; 2025-04-08
;
; https://www.autoitscript.com/forum/topic/191368-check-if-workstation-is-locked-cross-platform/
;
; =====================================================================

Func hr_IsWorkstationLocked()

    Local Const $WTS_CURRENT_SERVER_HANDLE = 0
    Local Const $WTS_CURRENT_SESSION = -1
    Local Const $WTS_SESSION_INFO_EX = 25

    Local $hWtsapi32dll = DllOpen("Wtsapi32.dll")

    Local $result = DllCall($hWtsapi32dll, "int", "WTSQuerySessionInformation", "int", $WTS_CURRENT_SERVER_HANDLE, "int", $WTS_CURRENT_SESSION, "int", $WTS_SESSION_INFO_EX, "ptr*", 0, "dword*", 0)

    If ((@error) OR ($result[0] == 0)) Then
        Return SetError(1, 0, False)
    EndIf

    Local $buffer_ptr = $result[4]
    Local $buffer_size = $result[5]

    Local $buffer = DllStructCreate("uint64 SessionId;uint64 SessionState;int SessionFlags;byte[" & $buffer_size - 20 & "]", $buffer_ptr)

    Local $isLocked = (DllStructGetData($buffer, "SessionFlags") == 0)

    $buffer = 0
    DllCall($hWtsapi32dll, "int", "WTSFreeMemory", "ptr", $buffer_ptr)
    DllClose($hWtsapi32dll)

    Return $isLocked

EndFunc