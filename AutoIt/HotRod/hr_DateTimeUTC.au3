#include-once

; =====================================================================
;	Return Current System Time as YYYY/MM/DD HH:MM UTC
; =====================================================================
	#include <Date.au3>

	Func hr_DateTimeUTC( )
	Local $ThisFunc = "hr_DateTimeUTC"

		Local $fmtYYYYsMMsDD = 1
		Local $dtCurrent = _Date_Time_GetSystemTime()

		Local $strCurrentDate = _Date_Time_SystemTimeToDateStr( $dtCurrent, $fmtYYYYsMMsDD )
		Local $strCurrentTime = StringTrimRight( _Date_Time_SystemTimeToTimeStr ( $dtCurrent ), 3 )
		Local $strCurrent = $strCurrentDate & " " & $strCurrentTime

		Return $strCurrent

	EndFunc

