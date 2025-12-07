Attribute VB_Name = "Globals_"
Option Explicit
Option Private Module

'         1         2         3         4         5         6         7         8         9
'123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890

' =====================================================================
'   Session Wide Global Constants and Variables
' =====================================================================
    
'   Known Views
'
Public Const glbProjects_PrimaryViewName    As String = "01 Projects"               ' Name of my Projects Folder Primary View
Public Const glbProjects_CatSearchViewName  As String = "TempCatSearch"             ' Name of my Projects Folder Temp Cat Search View

'   DLL Pointers
'
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)      ' Windows Sleep DLL Call

'   Response Types
'
Public Const glbResponse_Reply      As Integer = 1                                  ' Reply
Public Const glbResponse_ReplyAll   As Integer = 2                                  ' Reply All
Public Const glbResponse_Forward    As Integer = 3                                  ' Forward

'   Event_ExitEventScope - Where to go after Exit
'
Public Const glbExitEventScope_goNoWhere            As Long = 0                     ' The End
Public Const glbExitEventScope_goCustFormOpen       As Long = 1                     ' CustForm_Open

'   Error.Numbers
'
Public Const glbError_None                          As Long = 0                     ' No Error
Public Const glbError_InvalidProcArg                As Long = 5                     ' Invalid procedure call or argument (E_InvalidArg)
Public Const glbError_SubscriptOutOfRange           As Long = 9                     ' Just your garden variety array reference problem
Public Const glbError_TypeMismatch                  As Long = 13                    ' First seen in Misc_OLSetProperty using a Form property variable.
Public Const glbError_AppOrObjectDefinedError       As Long = 287                   ' Catch All of all catch alls
Public Const glbError_Automation430                 As Long = 430                   ' Catch All Method/Property not available
Public Const glbError_PropertyNotFound              As Long = 438                   ' Property does not exist. Object doesn't support this property or method.
Public Const glbError_Automation440                 As Long = 440                   ' Catch All Method/Property not available
Public Const glbError_ObjectNoActionSupport         As Long = 445                   ' Object doesn't support this action
Public Const glbError_CommandFailed                 As Long = 4198                  ' Object has been deleted
Public Const glbError_ObjectDeleted                 As Long = 5825                  ' Object has been deleted
Public Const glbError_MemberNotFound                As Long = 5941                  ' Member of the collection does not exist
Public Const glbError_Automation10409               As Long = -1040973551           ' While walking all stores
Public Const glbError_IMAP_NoUIDDelete              As Long = -2146644781           ' Specific version of glbError_CopiedNotMoved. Stupid is lost and sent an IMAP Delete with no UID (Server Responded 'Could not parse command')
Public Const glbError_ItemNotInCollection           As Long = -2147024809           ' Item not found by name in Collection
Public Const glbError_OpNotSupported                As Long = -2147024891           ' Item does not support this operation
Public Const glbError_DoNotHavePermissions          As Long = -2147024891           ' You do not have permissions to perform ... Can mean the item is Read-Only or other cases.
Public Const glbError_RegDeleteKeyNotFound          As Long = -2147024894           ' Registry Delete Key not found
Public Const glbError_RegReadKeyNotFound            As Long = -2147024894           ' Registry Read Key not found
Public Const glbError_CopiedNotMoved                As Long = -2147219840           ' Dana email error. Items were copied instead of moved because can't delete.
Public Const glbError_MsgInterfaceUnknown           As Long = -2147220991           ' Operation failed. Messaging interfaces unknown error.
Public Const glbError_CleanupHTML_Hoarked           As Long = -2147221221           ' Cleanup. Cleaned HTML won't go back into the Item. HTML too complex?
Public Const glbError_ObjectNotFound                As Long = -2147221233           ' General Object gone and Zombie IMAP Item Access (Item already deleted)
Public Const glbError_PAPropertyNotFound            As Long = -2147221233           ' PropertyAccessor. Property is unknown or cannot be found.
Public Const glbError_ItemIsZombie                  As Long = -2147221238           ' Zombie Item Access (Item already moved, deleted, sent, etc)
Public Const glbError_FormPageInvalid               As Long = -2147352567           ' Form Page does not exist or does not allow operation
Public Const glbError_ArayIndexOutOfBounds          As Long = -2147352567           ' Catch All index invalid
Public Const glbError_OpenRecurringFromInstance     As Long = -2147467259           ' Attempt to open a recurring meeting/appointment from an instance update

'   Error.Descriptions (For when Stupid uses a different Error.Number for the same error)
'
Public Const glbErrorDesc_PropertyIsReadOnly As String = "Property is read-only."
Public Const glbErrorDesc_TypeMismatch As String = "Type mismatch"
Public Const glbErrorDesc_ExplorerCannotBeUsed As String = "The Explorer has been closed and cannot be used for further operations. Review your code and restart Outlook."
Public Const glbErrorDesc_CannotSaveThisItem As String = "Cannot save this item."
Public Const glbErrorDesc_AlreadyDownloadedPrefix As String = "This item is already downloaded"

'   Private Errors
'
Public Const glbPrivateError_InvalidProcName As String = "0516 Invalid Proc Name"

'   idMSO Values
'
Public Const glbidMSO_AddressBook               As String = "AddressBook"               ' Open the default Address Book dialog
Public Const glbidMSO_AllCategories             As String = "AllCategories"             ' Open the All Categories (Master Cat) dialog
Public Const glbidMSO_ApplicationOptionsDialog  As String = "ApplicationOptionsDialog"  ' Used to tell if Explorer is fully active before doing an MSO command.
Public Const glbidMSO_AttachFile                As String = "AttachFile"                ' Used to tell if an item is in Edit Mode
Public Const glbidMSO_BulletsGalleryWord        As String = "BulletsGalleryWord"        ' Word - Toggle bullet formating
Public Const glbidMSO_Delete                    As String = "Delete"                    ' Menus -> Edit -> Delete
Public Const glbidMSO_EditMessage               As String = "EditMessage"               ' Put an item in Edit Mode
Public Const glbidMSO_FilePrint                 As String = "FilePrint"                 ' Open the File -> Print page
Public Const glbidMSO_FileSaveAs                As String = "FileSaveAs"                ' Used to tell if Inspector is fully active before doing an MSO command.
Public Const glbidMSO_FontDialog                As String = "FontDialog"                ' Word - Open the Font dialog
Public Const glbidMSO_Forward                   As String = "Forward"                   ' Forward button on Mail Items
Public Const glbidMSO_Reply                     As String = "Reply"                     ' Reply button on Mail Items
Public Const glbidMSO_ReplyAll                  As String = "ReplyAll"                  ' Reply All button on Mail Items
Public Const glbidMSO_SendDefault               As String = "SendDefault"               ' Send button on Mail Items
Public Const glbidMSO_SendReceiveAll            As String = "SendReceiveAll"            ' Send/Receive Send/Receive All Folders
Public Const glbidMSO_ProcessMarkedHeaders      As String = "SyncThisfolderMarked"      ' Send/Receive Process Marked Headers
Public Const glbidMSO_UpdateFolder              As String = "UpdateFolder"              ' Send/Receive Update Folder
Public Const glbidMSO_ViewVisualBasicCode       As String = "ViewVisualBasicCode"       ' Developer -> Form -> View Code

'   Standard MAPI Prop Tags
'
Private Const SchemasPrefix                     As String = "http://schemas.microsoft.com/mapi/"

Public Const glbPropTag_InviteSent              As String = SchemasPrefix & "id/{00062002-0000-0000-C000-000000000046}/8229000B"                ' PSETID_Meeting As Boolean
Public Const glbPropTag_IMAPStatus              As String = SchemasPrefix & "id/{00062008-0000-0000-C000-000000000046}/85700003"                ' "IMAP Status" "Marked For Deletion" As Boolean. No PidTag that I can find.
Public Const glbPropTag_InternetReplyID         As String = SchemasPrefix & "proptag/0x1042001F"                                                ' PR_IN_REPLY_TO_ID As String
Public Const glbPropTag_InternetMsgID           As String = SchemasPrefix & "proptag/0x1035001F"                                                ' PR_INTERNET_MESSAGE_ID_W As String
Public Const glbPropTag_FlagRequest             As String = "urn:schemas:httpmail:messageflag"                                                  ' FlagRequest. For my Post forms that don't expose it.
Public Const glbPropTag_IPFFolderType           As String = SchemasPrefix & "proptag/0x3613001E"                                                ' PR_CONTAINER_CLASS. Folder Type.
Public Const glbPropTag_SenderName              As String = SchemasPrefix & "proptag/0x0C1A001F"                                                ' PR_SENDER_NAME_W as String from mail. READ ONLY. And so is MailItem.SenderName.
Public Const glbPropTag_SentRepresenting        As String = SchemasPrefix & "proptag/0x0042001F"                                                ' PR_SENT_REPRESENTING_NAME_W as String from mail.
Public Const glbPropTag_Categories              As String = "urn:schemas-microsoft-com:office:office#Keywords"                                  ' Categories String
Public Const glbPropTag_CategoriesMAPI          As String = SchemasPrefix & "string/{00020329-0000-0000-C000-000000000046}/Keywords/0x0000101F" ' Categories MAPI
Public Const glbPropTag_MessageDeliveryTime     As String = SchemasPrefix & "proptag/0x0E060040"                                                ' PR_MESSAGE_DELIVERY_TIME as PT_SYSTIME
Public Const glbPropTag_BlockStatus             As String = SchemasPrefix & "proptag/0x10960003"                                                ' PR_BLOCK_STATUS as PT_I4 (Integer32)

'   User Props, Tags, and Values
'
Public Const glbUserPropsSchemaPrefix           As String = SchemasPrefix & "string/{00020329-0000-0000-C000-000000000046}/"                    ' PS_PUBLIC_STRINGS (UserProps)
Public Const glbUserPropsSchemaSuffix           As String = "/0x0000001F"                                                                       ' Property Type = PT_UNICODE).

'   HotRodGUID
'
'   ! Any changed here will require updating all linked Item and HyperLink Screen Tips !
'
Public Const glbUserPropTag_HotRodGUID          As String = "HotRodGUID"            ' HotRod GUID String.
Public Const glbUserPropTag_HotRodEntryId       As String = "HotRodEntryId"         ' EntryID As String. Value when HotRodGUID created or last updated.
Public Const glbUserPropTag_HotRodEntryIdMod    As String = "HotRodEntryIdMod"      ' Timestamp As String. Last time HotRodEntryId created or updated.
Public Const glbHotRodGUIDLabel                 As String = "HotRodGUID" & " "      ' HotRod GUID Label/Prefix (e.g. Screen Tip)

'   Custom Form Custom Actions
'

'   Global Class Instances
'
Public glbAppShadows    As AppShadows                                               ' Owner Class for InspShadow and ExplShadow instances
Public glbAppTimers     As AppTimers                                                ' Owner Class for Timer instances
Public glbProcXeq       As ProcXeq                                                  ' Single instance of the ProcXeq class


'   Reserved Unicode Characters
'
Public Const glbUnicode_RepSepMark      As Long = &H25B2                            ' U+25B2 - Black Up-Pointing Triangle
Public Const glbUnicode_RepSepMarkName  As String = "Responce Seperator"
Public Const glbUnicode_LineAnchor      As Long = &H25B3                            ' U+25B3 - White Up-Pointing Triangle
Public Const glbUnicode_LineAnchorName  As String = "Line Anchor"
Public Const glbUnicode_BQStartMark     As Long = &H25B6                            ' U+25B6 - Black Right-Pointing Triangle
Public Const glbUnicode_BQStartMarkName As String = "BlockQuote Start"
Public Const glbUnicode_BQEndMark       As Long = &H25C0                            ' U+25C0 - Black Left-Pointing Triangle
Public Const glbUnicode_BQEndMarkName   As String = "BlockQuote End"

'   Special Unicode Characters
'
Public Const glbUnicode_ZWNBSP          As Long = &HFEFF                            ' U+FEFF - Zero Width No-Break Space
Public Const glbUnicode_DownTriangle    As Long = &H25BC                            ' U+25BC - Black Down-Pointing Triangle (Tahoma)

'   Special Cat Prefixes (Values set in Global_Init)
'
Public glbSpecialCatPrefixTable         As String                                   ' Table of Special Cat Prefixes

Public glbCatPrefixNoCats               As String                                   ' Asterick
Public glbCatPrefixPriority             As String                                   ' Hyphen
Public glbCatPrefixFollowUp             As String                                   ' Triangle
Public glbCatPrefixTailEnd              As String                                   ' Omega (Sorts out last. Tail End Charlie).

'   WIP Cat
'
Public Const WIPCatPrefix               As String = "wip "                          ' WIP Master Cat prefix

'   Known Cats (Values set in Global_Init)
'
Public glbCatJunk                       As String
Public glbCatFollowUp                   As String
Public glbCatDeleted                    As String
Public glbCatNoCats                     As String
Public glbCatPriorityHigh               As String
Public glbCatPriorityMedium             As String
Public glbCatPriorityLow                As String
Public glbCatPriorityWait               As String

'   BillingInformation (BI) - When used to stash my stuff
'
Public Const glbBI_Sig  As String = "4B4415BE-A758-12B4-A021-9366A972ADAD"          ' GUID as first part of BI. Means it has my data in it.
Public Const glbBI_Sep  As String = vbVerticalTab                                   ' Part seperator in a BI String

Public Const glbBIInx_SIG           As Long = 0                                     ' SIG string index in BIData()
Public Const glbBIInx_Cats          As Long = 1                                     ' Cats string index in BIData()
Public Const glbBIInx_FlagRequest   As Long = 2                                     ' Follow Up FlagRequest string index in BIData()
Public Const glbBIInx_ReminderTime  As Long = 3                                     ' Follow Up ReminderTime string index in BIData()
Public Const glbBI_UBound           As Long = 3                                     ' Dim of BIData (Number of parts in a BI string -1)

'   Known Paths
'
'       Any chnages here must be reflected in Folders_KnownPaths
'
Public Const glbKnownPath_HomeInBox         As String = "\\GMail\Inbox"             ' Alais for GMail Inbox
Public Const glbKnownPath_Projects          As String = "\\Projects\Projects"       ' Current Projects
Public Const glbKnownPath_Inbox             As String = "\\Default\Inbox"           ' Default Inbox
Public Const glbKnownPath_Drafts            As String = "\\Default\Drafts"          ' Default Drafts
Public Const glbKnownPath_Contacts          As String = "\\Default\Contacts"        ' Default Contacts
Public Const glbKnownPath_GMailInbox        As String = "\\GMail\Inbox"             ' xxxx@gmail.com Inbox
Public Const glbKnownPath_Deleted           As String = "\\Default\Deleted Items"   ' Default Deleted Items
Public Const glbKnownPath_Journal           As String = "\\Default\Journal"         ' Default Journal
Public Const glbKnownPath_Outbox            As String = "\\Default\Outbox"          ' Default Outbox
Public Const glbKnownPath_ProjectsArchive   As String = "\\Projects - Archive - 2025\Projects - Archive - 2025"     ' Current Projects Archive

'   Junk Temp
'
Public Const glbJunkFilePath        As String = "C:\Junk\Temp\Outlook\"                 ' Path to the Junk Temp Outlook dir
Public Const glbViewsStartFolder    As String = glbJunkFilePath & "ViewSave"            ' Default starting folder for Views Save/Restore folder select
Public Const glbHotRodLogFile       As String = glbJunkFilePath & "Logs\HotRodLog.tsv"  ' FileSpec of the HotRod Log file
Public Const glbFileSaveHTML        As String = glbJunkFilePath & "HTML"                ' Dir for File_SaveHTML Files
Public Const glbTempFilePath        As String = glbJunkFilePath & "Temp\"               ' Temp File Directory Path (Attachments, File_SaveToTemp)
Public Const glbExportAsPDFPath     As String = glbJunkFilePath & "ExportAsPDF\"        ' Dir for Mail_ExportAsPDF

'   Master Cats List Backup
'
Public Const glbMasterCatsBackupFolder            As String = "C:\Data\Backups\Outlook\Categories"    ' Folder used by Categories_MasterCatsBackup
Public Const glbMasterCatsBackupVersionsToKeep    As Long = 255                                       ' Keepers for Categories_MasterCatsBackup

'   Custom Forms
'
Public Const glbCustForm_Card           As String = "IPM.Post.CardV3"
Public Const glbCustForm_WipProject     As String = "IPM.Post.WIPProjectV3"
Public Const glbCustForm_WipActivity    As String = "IPM.Post.WIPActivityV3"

Public Const gblCustFormType_Card       As String = "Card"
Public Const gblCustFormType_WIP        As String = "WIP"

'   Date and Time
'
Public glbDateNone                      As Date                                             ' VBA Date "None". = #4501-01-01#
Public glbDateMax                       As Date                                             ' Max Valid Date for Form fields (that I could find).
Public glbDateExpiresFlag               As Date                                             ' Special Expires flag for Custom Forms.

'   Misc
'
Public Const glbHotRodStyle             As String = "HotRod_Normal"                         ' Name of the Hot Rod Normal Style
Public glbCleanupAuto                   As Boolean                                          ' Automatic Cleanup on Response and Send?
Public Const glbMaxMessageSize          As Long = 20971520                                  ' Warn on send if message is over this size ( 20 * 1024 * 1024 )
Public glbLastResponseObject            As Object                                           ' Last Response object processed. See Response_Main.
Public Const glbMsg_MaxWidth            As Integer = 108                                    ' Max width of a Msg_ box in spaces.
Public Const glbBlankLine               As String = vbNewLine & vbNewLine                   ' Blank Line
Public Const glbQuote                   As String = """"                                    ' Double Quotes
Public Const glbViewsSaveGUID           As String = "2A75D2E4-757D-1217-AFE5-7A4A5146A336"  ' GUID in the 2nd line comment of my View save files.
Public Const glbGUIDLen                 As Long = 36                                        ' Len of all my GUID Strings
Public Const vbEmptyString              As String = vbNullString                            ' What it should have been called.
Public Const vbStrCompEqual             As Integer = 0                                      ' Because StrComp Equal is Zero. And I always think it's a mistake.
Public Const glbVBAProjectBackupKeepers As Long = 255                                       ' VBAProject backup older versions to keep. Used by File_BackupVBAProject_Trim
Public Const glbCatSep                  As String = ","                                     ' Cat String Seperator

' ---------------------------------------------------------------------
'   Shared with HotRod AutoIt
' ---------------------------------------------------------------------

'   HotRod External Programs
'
'       Can NOT build the program links as Const because of %HotRod% in the path.
'       Have to use as "glbHotRodLnks & glbHotRodLnks_XXXX"
'
Private Const glbHotRodEnvLiteral               As String = "HotRod"                        ' %HotRod% before it is expanded
Public glbHotRodEnv                             As String                                   ' %HotRod% after it is expanded

Public glbHotRodLnks                            As String                                   ' = "%HotRod%\LNK\" with Environment Variable expanded
Public Const glbHotRodLnks_CatDialogResize      As String = "Outlook_CatDialogResize.lnk"   ' LNK to Resize the Master Cats dialog AutoIt script
Public Const glbHotRodLnks_FileDialogs          As String = "File_FileDialogs.lnk"          ' LNK to the Standard File Dialogs AutoIt EXE
Public Const glbHotRodLnks_ExplorerOpenPath     As String = "Explorer_OpenPath.lnk"         ' LNK to Explorer Open Path AutoIt EXE

' ---------------------------------------------------------------------
'   Reg Keys
' ---------------------------------------------------------------------

'
'   Prefixes
'
Public Const glbRegHotRodPrefix                 As String = "HKEY_CURRENT_USER\SOFTWARE\VB and VBA Program Settings\HotRod\"
Public Const glbRegOutlookBaseKey               As String = glbRegHotRodPrefix & "Outlook\"

'   Outlook Shared Globals - Any changes must be reflected in C:\HotRod\AU3\_INC\Outlook\Globals.au3
'
Public Const glbVBAProjectBackupFolder          As String = "C:\Data\Backups\Outlook\VbaProject"    ' VBAProject backup folder used by File_BackupVBAProject and AltVBA
Public Const glbAltVBAEnv                       As String = "Outlook_AltVBA"                        ' Env Var Name. If defined - Started from my Outlook_AltVBA HotRod script. Value = VBAProject.OTM file path.

'   ProxXeq - Any changes must be reflected in C:\HotRod\AU3\_INC\Outlook\ProcXeq.au3
'
Public Const glbProcXeq_ActionName          As String = "ProcXeqAction"                     ' RSVP Journal Item Custom Action
Public Const glbProcXeq_RegBaseKey          As String = glbRegOutlookBaseKey & "ProcXeq\"   ' Base Key for ProcXeq Reg entries
Public Const glbProcXeq_RegStatus           As String = glbProcXeq_RegBaseKey & "Status"    ' Current Status
Public Const glbProcXeq_Status_Ready        As String = "Ready"                             ' Ready for new CmdLine to xeq
Public Const glbProcXeq_Status_Submitted    As String = "Submitted"                         ' CmdLine received
Public Const glbProcXeq_Status_Running      As String = "Running"                           ' CmdLine being processed
Public Const glbProcXeq_Status_Canceled     As String = "Canceled"                          ' AutoIt timed out waiting for me. Status <- Ready.
Public Const glbProcXeq_RegCmdLine          As String = glbProcXeq_RegBaseKey & "CmdLine"   ' CmdLine to Xeq
Public Const glbProcXeq_CmdLineSep          As Long = 9                                     ' CmdLine seperator {TAB}. Stupid doesn't allow Chr in Declarations Section.
Public Const glbProcXeq_RegEntryId          As String = glbProcXeq_RegBaseKey & "EntryID"   ' RSVP Journal Item Entry ID
Public Const glbProcXeq_RegStoreId          As String = glbProcXeq_RegBaseKey & "StoreID"   ' RSVP Journal Item Store ID

'   FileDialogs Shared Globals - Any changes must be reflected in C:\HotRod\AU3\_INC\File\FileDialogs\Globals.au3
'
Public Const glbFileDialogs_Action_FileOpen     As String = "FileOpen"
Public Const glbFileDialogs_Action_FileSave     As String = "FileSave"
Public Const glbFileDialogs_Action_FolderSelect As String = "FolderSelect"
Public Const glbFileDialogs_RegBaseKey          As String = glbRegHotRodPrefix & "FileDialogs\"         ' Base Key for FileDialogs AutoIt Script
Public Const glbFileDialogs_RegCmdLine          As String = glbFileDialogs_RegBaseKey & "CmdLine"       ' -> "FileOpen", "FileSave", "FolderSelect" + args
Public Const glbFileDialogs_CmdLineSep          As Long = 9                                             ' CmdLine seperator {TAB}. Stupid doesn't allow Chr in Declarations Section.
Public Const glbFileDialogs_RegRetunKey         As String = glbFileDialogs_RegBaseKey & "Return\"       ' Base Key for Return values
Public Const glbFileDialogs_RegResults          As String = glbFileDialogs_RegRetunKey & "Results"      ' <- FileSpec(s) or Folder
Public Const glbFileDialogs_RegCanceled         As String = glbFileDialogs_RegRetunKey & "Canceled"     ' "0", "1". (Also Msg and Non-Zero Exit Code on Error)
    
' ---------------------------------------------------------------------
'   Shells
' ---------------------------------------------------------------------
'

'   Reference: Windows Script Host Object Model, IWshRuntimeLibrary, C:\Windows\SysWOW64\wshom.ocx
'   AKA: Windows Scripting Shell, Windows Scripting Host, Windows Script Host, WSH, "WScript.Shell", IWshRuntimeLibrary.WshShell
'
'       https://ss64.com/vb/shell.html
'       https://ss64.com/vb/run.html
'
'       .AppActivate            Activate a window by Title
'       .Run                    Run an application.
'       .RegRead/Delete/Write   Registry Operations
'
Public glbWshShell As IWshRuntimeLibrary.WshShell

'   Reference: Microsoft Shell Controls And Automation, Shell32, C:\Windows\SysWOW64\shell32.dll
'   AKA: Windows Application Shell, "Shell.Application", Shell32.Shell
'
'       https://ss64.com/vb/application.html
'       https://ss64.com/vb/shellexecute.html
'       https://learn.microsoft.com/en-us/windows/win32/shell/shell#methods
'       https://www.devhut.net/vba-shell-application-deep-dive/
'
'       .ExpandEnvironmentStrings   Expand a Windows environment variable.
'       .ShellExecute               Run a script or application.
'       {And LOTS more I don't use}
'
Public glbAppShell As Shell32.Shell

'   Reference: Microsoft Scripting Runtime, Scripting, C:\Windows\SysWOW64\scrrun.dll
'   AKA: Windows Scripting
'
'       .Dictionary
'
'           https://excelmacromastery.com/vba-dictionary/
'
'       File System Object
'
'           https://ss64.com/vb/filesystemobject.html
'           https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/filesystemobject-object#methods
'           https://www.thevbprogrammer.com/ch06/06-09-fso.htm
'
'           {Anything to do with files}
'
Public glbFSO As Scripting.FileSystemObject                                                         '
    
' =====================================================================
'   Init my Globals
' =====================================================================

Public Function Globals_Init() As Boolean
Const ThisProc = "Globals_Init"
Globals_Init = False

    '   Windows Scripting objects
    '
    If glbWshShell Is Nothing Then Set glbWshShell = New IWshRuntimeLibrary.WshShell
    If glbAppShell Is Nothing Then Set glbAppShell = New Shell32.Shell
    If glbFSO Is Nothing Then Set glbFSO = New Scripting.FileSystemObject
    
    '   HotRod Environment Var
    '
    If Not Misc_EnvironmentGet(glbHotRodEnvLiteral, glbHotRodEnv) Then Stop: Exit Function
    glbHotRodLnks = glbHotRodEnv & "\LNK\"
    
    '   Automatic Cleanup?
    '
    glbCleanupAuto = True
    
    '   TempFilePath
    '
    If Not glbFSO.FolderExists(glbTempFilePath) Then
        Msg_Box Text:="Cannot find the TempFile folder '" & glbTempFilePath & "'. Please create it and Restart Outlook.", Proc:=ThisProc
        Exit Function
    End If
    
    '   See Response_Main for what this is for
    '
    Set glbLastResponseObject = Nothing
    
    '   Special Cat Prefixes
    '
    glbCatPrefixNoCats = "*"
    glbCatPrefixPriority = "-"
    glbCatPrefixFollowUp = ChrW(&H25B6)         ' U+25B6 Black Right-Pointing Triangle
    glbCatPrefixTailEnd = ChrW(&H3A9)           ' U+03A9 Greek Capital Letter Omega
    
    glbSpecialCatPrefixTable = _
    "'                                          " & vbLf & _
    "  Prefix                      | Name       " & vbLf & _
    "' --------------------------- | -----------" & vbLf & _
    " " & glbCatPrefixNoCats & "   | No Cats    " & vbLf & _
    " " & glbCatPrefixPriority & " | Priority   " & vbLf & _
    " " & glbCatPrefixFollowUp & " | Follow Up  " & vbLf & _
    " " & glbCatPrefixTailEnd & "  | Tail End   " & vbLf & _
    "'"
    
    '   Known Cats
    '
    '       (See also CustForm_WipProjWrite AssignCats:)
    '
    glbCatJunk = "Junk"
    glbCatFollowUp = glbCatPrefixFollowUp & " Follow Up"
    glbCatDeleted = glbCatPrefixTailEnd & " Deleted"
    glbCatPriorityHigh = glbCatPrefixPriority & " High"
    glbCatPriorityMedium = glbCatPrefixPriority & " Medium"
    glbCatPriorityLow = glbCatPrefixPriority & " Low"
    glbCatPriorityWait = glbCatPrefixPriority & " Wait"
    glbCatNoCats = glbCatPrefixNoCats & " No Cats"
    
    '   Init my Shadows
    '
    Set glbAppShadows = Nothing
    Set glbAppShadows = New AppShadows
    If Not glbAppShadows.Initialize Then Stop: Exit Function
    
    '   Init Misc Events Hooks
    '
    If Not ThisOutlookSession.MiscEventsHook Then Stop: Exit Function

    '   Init my Timers
    '
    Set glbAppTimers = Nothing
    Set glbAppTimers = New AppTimers
    If Not glbAppTimers.Initialize Then Stop: Exit Function
    
    '   ProcXeq Init
    '
    Set glbProcXeq = Nothing
    Set glbProcXeq = New ProcXeq
    If Not glbProcXeq.Initialize Then Stop: Exit Function
    
    '   Dates
    '
    glbDateNone = DateSerial(4501, 1, 1)
    glbDateMax = DateSerial(4499, 12, 31)
    glbDateExpiresFlag = DateSerial(3999, 9, 9) + TimeSerial(9, 9, 9)
    
Globals_Init = True
End Function


