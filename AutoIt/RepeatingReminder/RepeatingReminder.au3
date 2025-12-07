#AutoIt3Wrapper_OutFile_Type=EXE
#AutoIt3Wrapper_Icon=RepeatingReminder.ico
#include "HotRod\hr_Directives.au3"

; =====================================================================

;  Title Match Mode = any substring in the title
	Opt("WinTitleMatchMode", 2)

;	Show a Tray Icon
	Opt("TrayIconHide", 0)
;
; Globals - External Pointers
;

	Global $Chime_Wave = "C:\Data\Exchange\Common\Waves\Big Ben Single Chime.wav"

	; 2012-01-20 Changed from "wv_player.exe" to "playwav.exe" for wav player
	; because wv_player was waking up the machine from idle state
	; and turning on the monitors, that had been powered off

	Global $Wave_PlayerExe = "C:\Prog\Acc\Sound\PlayWav\playwav.exe"

; Globals - Internal Use

	Global $TimerRunning = False
	Global $TimerStart

	Global Const $PlayDelay = ( 60 * 1000 )
	Global Const $LoopDelay = 1000

; Main

	While True

		ConsoleWrite( "TimerDiff( $TimerStart ) = '" & TimerDiff( $TimerStart ) & "'." & @CRLF )
		ConsoleWrite( "$PlayDelay = '" & $PlayDelay & "'." & @CRLF )

		Select
			Case ReminderWindowExist() = False
				ConsoleWrite( "ReminderWindowExist() = False" & @CRLF )
				$TimerRunning = False
			Case $TimerRunning = False
				ConsoleWrite( "$TimerRunning = False" & @CRLF )
				Sound_Play( $Chime_Wave )
				$TimerRunning = True
				$TimerStart = TimerInit()
			Case TimerDiff( $TimerStart ) > $PlayDelay
				ConsoleWrite( "TimerDiff( $TimerStart ) > $PlayDelay" & @CRLF )
				$TimerRunning = False
		EndSelect

		Sleep( $LoopDelay )

	WEnd

Func ReminderWindowExist()

	Local $hWnd = WinGetHandle("[TITLE:Reminder; CLASS:#32770]", "Click Snooze to be reminded again")
	; ConsoleWrite( "$hWnd = '" & $hWnd & "'." & @CRLF )
	If @error Then Return False

	Local $hCtrl = ControlGetHandle($hWnd, "", "[CLASS:SysListView32; INSTANCE:1; ID:8342]")
	; ConsoleWrite( "$hCtrl = '" & $hCtrl & "'."  & @CRLF )
	If @error Then Return False

	Return True

EndFunc

; ---------------------------------------------------------------------
; 2012-01-11 - Chnaged to use external program to play the sound
; because internal AutoIT SoundPlay and _SoundXXXXX UDF could fail without error sometimes

Func Sound_Play( $SoundFileName )

	ConsoleWrite( "Sound_Play Called" & @CRLF )

	If Not FileExists( $SoundFileName ) Then
		MsgBox(16,"Outlook Repeating Reminder - Sound_Play - Error","Sound File '" & $SoundFileName & "' does not exist.")
		Exit
	EndIf
	If Not FileExists( $Wave_PlayerExe ) Then
		MsgBox(16,"Outlook Repeating Reminder - Sound_Play - Error","Wave Player exe '" & $Wave_PlayerExe & "' does not exist.")
		Exit
	EndIf

	ShellExecuteWait( $Wave_PlayerExe, '"' & $SoundFileName & '"', "", "", @SW_HIDE )
	If @error Then
		MsgBox(16,"Outlook Repeating Reminder - Sound_Play - Error","ShellExecuteWait error calling Wave Player '" & $Wave_PlayerExe & "' for file '" & $SoundFileName & "'.")
		Exit
	EndIf

EndFunc
