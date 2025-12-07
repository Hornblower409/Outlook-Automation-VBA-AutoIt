#include-once

; Outlook Globals

	;	Shared with my Outlook VBA. Any changes should be reflected in VBA Globals_ module.

		; 2025-01-14 - Unused
		; Temp File Directory Path
		; Global Const $Outlook_Global_TempFilePath = "C:\Data\Exchange\TEMP\"

;	2024-11-11 - Unused in AutoIt. Switched to having C:\HotRod\AU3\Outlook\ProjectsCatCheck
;	use ProcXeq to call "PROJECTS_ITEMADDEXTERNAL" so Outlook Macro does the processing.
;
;~ 		; BillingInformation field sig. Means it has Cats in Hex (WITH Trailing Space)
;~ 		Global Const $Outlook_Global_BillingInfoSig = "4B4415BE-A758-12B4-A021-9366A972ADAD "

		; VBAProject backup folder used by File_BackupVBAProject and AltVBA
		Global Const $Outlook_Global_VBAProjectBackupFolder = "C:\Data\Backups\Outlook\VbaProject"

		; Env Var Name. If defined - Started from my Outlook_AltVBA HotRod script. Value = VBAProject.OTM file path.
		Global Const $Outlook_Global_AltVBAEnv = "Outlook_AltVBA"

		; Base Key for all of my Registry entries
		Global Const $Outlook_gblRegHotRodPrefix = "HKEY_CURRENT_USER\SOFTWARE\VB and VBA Program Settings\HotRod\"
		Global Const $Outlook_gblRegBaseKey = $Outlook_gblRegHotRodPrefix & "Outlook\"

		; Global Const $Outlook_gblRegBaseKey = "HKEY_CURRENT_USER\SOFTWARE\VB and VBA Program Settings\"

		;   ProxXeq
		;
			Global Const $Outlook_gblProcXeq_ActionName = "ProcXeqAction"											; RSVP Journal Item Custom Action
			Global Const $Outlook_gblProcXeq_RegBaseKey = $Outlook_gblRegBaseKey & "ProcXeq\"						; Base Key for ProcXeq Reg entries
			Global Const $Outlook_gblProcXeq_RegStatus = $Outlook_gblProcXeq_RegBaseKey & "Status"					; Current Status
			Global Const $Outlook_gblProcXeq_Status_Ready = "Ready"													;   Ready for new CmdLine to xeq
			Global Const $Outlook_gblProcXeq_Status_Submitted = "Submitted"											;   CmdLine received
			Global Const $Outlook_gblProcXeq_Status_Running = "Running"												;   CmdLine being processed
			Global Const $Outlook_gblProcXeq_Status_Canceled = "Canceled"											;   AutoIt timed out waiting for Outlook. Status <- Ready.
			Global Const $Outlook_gblProcXeq_RegCmdLine = $Outlook_gblProcXeq_RegBaseKey & "CmdLine"				; CmdLine to Xeq
			Global Const $Outlook_gblProcXeq_CmdLineSep = Chr( 9 )													; CmdLine seperator {TAB}
			Global Const $Outlook_gblProcXeq_RegEntryId = $Outlook_gblProcXeq_RegBaseKey & "EntryID"				; RSVP Journal Item Entry ID
			Global Const $Outlook_gblProcXeq_RegStoreId = $Outlook_gblProcXeq_RegBaseKey & "StoreID"				; RSVP Journal Item Store ID

	;	Outlook Standards

	Global Const $Outlook_olMSGUnicode = 9	; OlSaveAsType = Unicode message format (.msg)
	Global Const $Outlook_msoControlButton = 1 ; MsoControlType msoControlButton

