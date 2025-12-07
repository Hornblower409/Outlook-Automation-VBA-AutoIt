#include-once

; AWDs for the Initial Outlook Explorers

	Global Const $Inbox_WinTitle 	= "[TITLE:Inbox - GMail - Microsoft Outlook; CLASS:rctrl_renwnd32]"
	Global Const $Projects_WinTitle = "[TITLE:Projects - Projects - Microsoft Outlook; CLASS:rctrl_renwnd32]"
	Global Const $Cards_WinTitle 	= "[TITLE:Cards - Cards - Microsoft Outlook; CLASS:rctrl_renwnd32]"
	Global Const $WIP_WinTitle 		= "[TITLE:Projects - Projects - Microsoft Outlook; CLASS:rctrl_renwnd32]"
	Global Const $Calendar_WinTitle = "[TITLE:Calendar - Default - Microsoft Outlook; CLASS:rctrl_renwnd32]"

	;  - Only changes the title on the Taskbar. Not on the actual window.
	;  - But that's better than nothing!

	Global Const $Inbox_ShortName 		= "Inbox - Outlook"
	Global Const $Projects_ShortName 	= "Projects - Outlook"
	Global Const $Cards_ShortName 		= "Cards - Outlook"
	Global Const $WIP_ShortName 		= "WIP - Outlook"
	Global Const $Calendar_ShortName 	= "Calendar - Outlook"

; AWDs for the Outlook Explorers after I have change the titles

	Global Const $Inbox_ShortTitle 		= "[TITLE:" & $Inbox_ShortName 		& "; CLASS:rctrl_renwnd32]"
	Global Const $Projects_ShortTitle 	= "[TITLE:" & $Projects_ShortName	& "; CLASS:rctrl_renwnd32]"
	Global Const $Cards_ShortTitle 		= "[TITLE:" & $Cards_ShortName 		& "; CLASS:rctrl_renwnd32]"
	Global Const $WIP_ShortTitle 		= "[TITLE:" & $WIP_ShortName 		& "; CLASS:rctrl_renwnd32]"
	Global Const $Calendar_ShortTitle 	= "[TITLE:" & $Calendar_ShortName 	& "; CLASS:rctrl_renwnd32]"

; My Standard Explorer Window Position & Size

	Global Const $Std_X = 935
	Global Const $Std_Y = 0
	Global Const $Std_W = 805
	Global Const $Std_H = 1080

; Timers and Retries

	Global Const $win_Activate_Retries = 15
	Global Const $win_WaitExist_Timeout = 120
	Global Const $win_WaitActive_Timeout = 5
