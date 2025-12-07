#include "HotRod\hr_Directives.au3"
#include "Outlook\ProcXeq.au3"
#NoTrayIcon

; =====================================================================
; 2024-12-22
;
;	Execute an Outlook VBA Proc with Param from the Command Line
;
; 	%HotRod%\LNK\Outlook_ProcXeq.lnk  {/Options}  Command  {Args}
;
; =====================================================================
;
Main()
Func Main()
Local Const $ThisFunc = "Main"

	Outlook_ProcXeq( $CmdLine )
	Exit( 0 )

EndFunc
