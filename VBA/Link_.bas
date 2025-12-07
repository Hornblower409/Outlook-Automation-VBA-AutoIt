Attribute VB_Name = "Link_"
Option Explicit
' NOT Option Private Module - I want these Subs to be exposed

' ---------------------------------------------------------------------
'   Links - Public Sub entry points -> Internal Procs
'
'       So I can disconnect "External" macro references from my internal Proc names.
'
'       Some of the _LINK names are used for Ribbon commands. So if you change
'       any _Link names you may break an existing Ribbon reference.
'
'       Almost all my Modules have "Option Private Module" which means that any Subs
'       will NOT show up on the Developer -> Macros list. Even if they meet the all
'       the criteria for being a runable macro.
'
'       If you want something on the Macros list, put a link to it here.
'
'       No need to keep in sorted order. He'll do that when you dropdown a Macros list
'
' ---------------------------------------------------------------------
'
Public Sub Application_Init_Link():                             Event_InitApplication:                              End Sub
Public Sub Backup_VBAProject_Link():                            File_BackupVBAProject_Manual:                       End Sub
Public Sub Calendar_MeetingFromInvite_Link():                   Misc_CalendarOpenMeetingFromInvite:                 End Sub
Public Sub Cleanup_Inspector_Link():                            Cleanup_Manual:                                     End Sub
Public Sub Cleanup_Toggle_Link():                               Cleanup_Toggle:                                     End Sub
Public Sub Item_Inspect_Link():                                 Misc_ItemInspect:                                   End Sub
Public Sub Get_Link_Link():                                     HypLnk_HyperlinkGet:                                End Sub
Public Sub Save_InspectorHTML_ToFile_Link():                    File_SaveInspectorHTML:                             End Sub
Public Sub SmartDel_PurgeAll_Link():                            IMAP_SmartDelPurgeAll:                              End Sub
Public Sub Views_SaveAll_Link():                                Views_SaveAll:                                      End Sub
Public Sub WipProjNew_Link():                                   CustForm_WipProjNew:                                End Sub

' Public Sub Categories_ExportToFile_Link():                      Categories_ExportToFile:                            End Sub
' Public Sub HTML_Edit_Link():                                    HtmlEdit_Edit:                                      End Sub
' Public Sub Clone_Card_Link():                                   Cards_Clone:                                        End Sub
' Public Sub SnoozeBeforeStart_OneHour_Link():                    Reminder_SnoozeBeforeStart_OneHour:                 End Sub
' Public Sub SnoozeBeforeStart_OneMinute_Link():                  Reminder_SnoozeBeforeStart_OneMinute:               End Sub
' Public Sub ClearAddresses_Link():                               Utility_ClearAddresses:                             End Sub
' Public Sub Views_HSHoark_Link():                                Views_HSHoark:                                      End Sub
' Public Sub Views_LockAll_Link():                                Views_LockAll:                                      End Sub
' Public Sub Views_UnlockAll_Link():                              Views_UnlockAll:                                    End Sub
' Public Sub Views_ShowLock_Link():                               Views_ShowLock:                                     End Sub
' Public Sub Views_LockView_Link():                               Views_LockView:                                     End Sub
' Public Sub Views_UnlockView_Link():                             Views_UnlockView:                                   End Sub
' Public Sub CreateSendReceiveApplicationFoldersGroup_Link():     Utility_CreateSendReceiveApplicationFoldersGroup:   End Sub

