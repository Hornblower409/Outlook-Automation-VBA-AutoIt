Attribute VB_Name = "IMAP_"
Option Explicit
Option Private Module

' =====================================================================
'   Is {something| IMAP?
'
'       Just a bunch of different ways to get to IMAP_FolderIsIMAP
'
' =====================================================================

'   Is the Active Window open on an IMAP Folder?
'
Public Function IMAP_ActiveWindowIsIMAP() As Boolean

    IMAP_ActiveWindowIsIMAP = IMAP_WindowIsIMAP(ActiveWindow)
    
End Function

'   Is a Window open on an IMAP Folder?
'
Public Function IMAP_WindowIsIMAP(ByVal Window As Object) As Boolean
IMAP_WindowIsIMAP = False

    Dim Folder As Outlook.Folder
    If Not Folders_Window(Window, Folder) Then Exit Function
    IMAP_WindowIsIMAP = IMAP_FolderIsIMAP(Folder)

End Function

'   Is an Object from an IMAP Folder?
'
Public Function IMAP_ItemIsIMAP(ByVal Item As Object) As Boolean
IMAP_ItemIsIMAP = False

    Dim Folder As Outlook.Folder
    If Not Folders_Item(Item, Folder) Then Exit Function
    IMAP_ItemIsIMAP = IMAP_FolderIsIMAP(Folder)

End Function

'   Is a Folder IMAP?
'
Public Function IMAP_FolderIsIMAP(ByVal oFolder As Outlook.Folder) As Boolean

    IMAP_FolderIsIMAP = Folders_TypeIsIMAP(oFolder)

End Function

'   Is a Folder Path IMAP?
'
Public Function IMAP_FolderPathIsIMAP(ByVal FolderPath As String) As Boolean
IMAP_FolderPathIsIMAP = False

    Dim oFolder As Outlook.Folder
    Set oFolder = Folders_Path(FolderPath)
    If oFolder Is Nothing Then Exit Function
    IMAP_FolderPathIsIMAP = IMAP_FolderIsIMAP(oFolder)

End Function

' =====================================================================
'   InternetMsgID (MsgId)
' =====================================================================

'   Generate a GUID InternetMsgID
'
Public Function IMAP_NewInternetMsgID() As String

    IMAP_NewInternetMsgID = "<" & Misc_MakeGUID() & "@aglassman.gmail.com>"
    
End Function

'   Get the InternetMsgID of a Mail Item
'
Public Function IMAP_MsgID(ByVal Item As Object) As String

    If Not Misc_OLGetProperty(Item, glbPropTag_InternetMsgID, IMAP_MsgID) Then IMAP_MsgID = ""
    
End Function

'   Get all Items in a Folder with a specific InternetMsgID
'
'   SPOS - You can NOT do a DASL/JET Filter/Find/Restrict on ConversationIndex or ConversationID
'   because they are binary fields. So lets try this.
'
Public Function IMAP_FilterByMsgID(ByVal MsgId As String, ByVal Folder As Outlook.Folder, ByRef Results As VBA.Collection) As Boolean
IMAP_FilterByMsgID = False
    
    If MsgId = "" Then Exit Function

    '   Construct filter for InternetMsgID
    '
    Dim strFilter As String
    strFilter = "@SQL=" & glbQuote & glbPropTag_InternetMsgID & glbQuote & " = " & "'" & MsgId & "'"
    
    '   Get the Results VBA.Collection
    '
    If Not Collection_FromRestrict(strFilter, Folder, Results) Then Stop: Exit Function
    If Results.count > 0 Then IMAP_FilterByMsgID = True
     
End Function

Public Function IMAP_DeleteByMsgId(ByVal MsgId As String, ByVal Folder As Outlook.Folder) As Boolean
Const ThisProc = "IMAP_DeleteByMsgId"
IMAP_DeleteByMsgId = True

    '   Get a VBA.Collection of items with this MsgId
    '   If nothing found - Return True.
    '
    Dim Results As VBA.Collection
    If Not IMAP_FilterByMsgID(MsgId, Folder, Results) Then Exit Function
    
    '   Delete all Items in Results
    '
    Dim oItem As Object
    For Each oItem In Results
        oItem.Delete
    Next oItem
    
    '   If all gone - Return True
    '
    If Not IMAP_FilterByMsgID(MsgId, Folder, Results) Then Exit Function
    
IMAP_DeleteByMsgId = False
End Function

' =====================================================================
'   IMAP Status ("Marked For Deletion" flag)
' =====================================================================

'   Get the IMAP Status (aka "Marked For Deletion") flag for an IMAP item.
'
'       Appears to be Read Only as OLSetProperty looks like he changes it on the current item
'       but it doesn't seem to stick when you save. Not sure. Might be how I'm doing the Set.
'
Public Function IMAP_IMAPStatus(ByVal Item As Outlook.MailItem) As Boolean
    
    If Not Misc_OLGetProperty(Item, glbPropTag_IMAPStatus, IMAP_IMAPStatus) Then IMAP_IMAPStatus = False
    
End Function

' =====================================================================
'   IMAP Update/Download Folder
' =====================================================================

'   Do an Update Folder on a Known Path Folder (Silent Failure)
'
'   2024-12-14 - Update Folder only checks for new mail and downloads Headers.
'   Does NOT download full items no matter what is setup in Send/Receive Groups.
'
'   2025-03-09 - Called from Scheduled Task Outlook_UpdateFolder
'
Public Function IMAP_UpdateFolder(ByVal KnownPath As String) As Boolean
Const ThisProc = "IMAP_UpdateFolder"
IMAP_UpdateFolder = False

    '   Find an open (but maybe not Active) Explorer for the KnownPath
    '
    Dim IMAPExplorer As Outlook.Explorer
    Set IMAPExplorer = Folders_FolderExplorer(Folders_KnownPath(KnownPath))
    
    '   Silent Fail if
    '
    '   - No Explorer open on the KnownPath Folder.
    '   - No Update Folder button. (Explorer not in Normal state. e.g. Task List)
    
    If IMAPExplorer Is Nothing Then Exit Function
    If Not Ribbon_ExecuteMSO(IMAPExplorer, glbidMSO_UpdateFolder) Then Exit Function
   
IMAP_UpdateFolder = True
End Function

'   Download all items in an IMAP Folder
'
Public Function IMAP_DownloadFolder(ByVal KnownPath As String) As Boolean
Const ThisProc = "IMAP_DownloadFolder"
IMAP_DownloadFolder = False

    Dim IMAPFolder As Outlook.Folder
    Set IMAPFolder = Folders_KnownPath(KnownPath)

    '   Find an open (but maybe not Active) Explorer for the KnownPath
    '
    Dim IMAPExplorer As Outlook.Explorer
    Set IMAPExplorer = Folders_FolderExplorer(IMAPFolder)
    If IMAPExplorer Is Nothing Then Stop: Exit Function
    
    '   Do an Update Folder to get any new Headers
    '
    If Not Ribbon_ExecuteMSO(IMAPExplorer, glbidMSO_UpdateFolder) Then Stop: Exit Function
    
    '   Set All = MarkForDownload
    '
    '   - Because I am NOT going to fight with Header Status (aka HS, DownloadState)
    '   - See Mail_SearchDownloadState.
    '
    '   2025-05-02 - Added Ignore CatDeleted
    '   2025-05-03 - Added skip of Count = 0
    '
    Dim oItem As Object
    Dim MarkedCount As Long
    For Each oItem In IMAPFolder.Items
    
        If TypeOf oItem Is Outlook.MailItem Then
            If Not (oItem.Categories = glbCatDeleted) Then
            
                On Error Resume Next
                
                    oItem.MarkForDownload = Outlook.olMarkedForDownload
                    Select Case Err.Number
                        
                        '   Ready for Download
                        '
                        Case glbError_None
                            MarkedCount = MarkedCount + 1
                        
                        '   Already Downloaded
                        '
                        Case glbError_DoNotHavePermissions
                            If Not (InStr(1, Err.Description, glbErrorDesc_AlreadyDownloadedPrefix) = 1) Then Stop: Exit Function
                            
                        Case Else
                            Stop: Exit Function
                            
                    End Select
                    
                On Error GoTo 0
                
            End If
        End If
    
    Next oItem
    
    '   Process Marked Headers
    '
    If MarkedCount > 0 Then
        If Not Ribbon_ExecuteMSO(IMAPExplorer, glbidMSO_ProcessMarkedHeaders) Then Stop: Exit Function
    End If
    
IMAP_DownloadFolder = True
End Function

'   Get an IMAP Inbox Folder using Session.Folders
'
'   - Can NOT use normal Folders_Path because it sometimes causes a "Downloading folder order".
'   - This works. Don't know why. Don't care. I found it by trial and error.
'   - See Card: Downloading folder order (In Status Bar) | IMAP | Outlook | Software
'
Public Function IMAP_SessionInbox(ByVal FolderPath As String, ByVal SessionFolderName As String) As Outlook.Folder

    On Error GoTo ProcExit
    
        Dim oFolder As Outlook.Folder
        Set oFolder = Session.Folders.Item(SessionFolderName)
        Dim oInboxFolder As Outlook.Folder
        Set oInboxFolder = oFolder.Store.GetDefaultFolder(Outlook.olFolderInbox)
        
    On Error GoTo 0
    Set IMAP_SessionInbox = oInboxFolder
        
ProcExit:

    If IMAP_SessionInbox Is Nothing Then
        Debug.Print "Failed to get an Inbox Folder for SessionFolderName: " & SessionFolderName
        Stop: Exit Function
    End If

End Function

' =====================================================================
'   SmartDel - If IMAP then Assign Deleted Cat else Delete
' =====================================================================

Public Function IMAP_SmartDel() As Boolean
IMAP_SmartDel = False
    
    Dim VBACollection As VBA.Collection
    Set VBACollection = New VBA.Collection
    Dim Index As Long

    '   Build VBACollection from an Inspector or Explorer Selection
    '
    Select Case True
        Case TypeOf ActiveWindow Is Outlook.Inspector
            VBACollection.Add ActiveInspector.CurrentItem, CStr(1)
        Case TypeOf ActiveWindow Is Outlook.Explorer
            With ActiveExplorer.Selection
                For Index = 1 To .count
                    VBACollection.Add .Item(Index), CStr(Index)
                Next Index
            End With
        Case Else
            Stop: Exit Function
    End Select
    
    '   Walk the VBACollection
    '
    Dim oItem As Object
    For Index = 1 To VBACollection.count: Do
    
        Set oItem = VBACollection.Item(Index)
        
        '   If it's already gone - done
        '   If it will take a DelCat - done
        '
        If oItem Is Nothing Then Exit Do
        If IMAP_SmartDelCat(oItem) Then Exit Do
        
        '   Else Delete
        '
        On Error Resume Next
            oItem.Delete
            Select Case Err.Number
                Case glbError_None, glbError_AppOrObjectDefinedError
                    ' Continue
                Case Else
                    Stop: Exit Function
            End Select
        On Error GoTo 0
        
    Loop While False: Next Index

IMAP_SmartDel = True
End Function

Public Function IMAP_SmartDelCat(ByVal oItem As Object) As Boolean
Const ThisProc = "IMAP_SmartDelCat"
IMAP_SmartDelCat = False

    If Not TypeName(oItem) = "MailItem" Then Exit Function
    If Not IMAP_ItemIsIMAP(oItem) Then Exit Function
    If Not oItem.Saved Then Exit Function
    If Not oItem.Sent Then Exit Function

    '   2025-05-14 - Try update via PropAccessor to avoid server sync
    '   2025-06-03 - Doesn't seem to make any difference.
    '
    oItem.Categories = glbCatDeleted
    
    '   2025-09-27 - Still hangs on some Saves. Collect and Retry.
    '
Retry:

    On Error Resume Next
    
        oItem.Save
        Select Case Err.Number
        
            Case glbError_None
                ' Continue
            
            Case Else
            
                Select Case Msg_Box( _
                    oErr:=Err, _
                    Proc:=ThisProc, Step:="oItem.Save", _
                    Icon:=vbQuestion, Buttons:=vbYesNoCancel, Default:=vbDefaultButton1, _
                    Subject:=oItem.Subject, _
                    Text:="IMAP Item Save Failed." & glbBlankLine & _
                    "Retry?    (Cancel = Debug)")
                    
                    Case vbYes
                        GoTo Retry
                    Case vbNo
                        Exit Function
                    Case vbCancel
                        Stop
               
                End Select
                
        End Select
    
    On Error GoTo 0
   
    '   2025-06-26 - WTF?
    '
    '   oItem.Close works, even if the Item being operated on doesn't have an Inspector.
    '   i.e. A SmartDel on an IMAP Explorer Selection works. Even though the docs make
    '   it sound like it shouldn't.
    '
    '   https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.close(method)
    '   "To run this example, you need to have an item displayed in an inspector window"
    '
    oItem.Close Outlook.OlInspectorClose.olDiscard

IMAP_SmartDelCat = True
End Function

Public Function IMAP_SmartDelPurge(ByVal FolderPath As String) As Boolean
IMAP_SmartDelPurge = False

    Dim oFolder As Outlook.Folder
    Set oFolder = Folders_KnownPath(FolderPath)
    
    '   Get a Collection of Items with CatDeleted
    '   If nothing found - Return True.
    '
    Dim strFilter As String
    strFilter = "@SQL=" & glbQuote & glbPropTag_Categories & glbQuote & " = " & "'" & glbCatDeleted & "'"
    Dim Results As VBA.Collection
    If Not Collection_FromRestrict(strFilter, oFolder, Results) Then Stop: Exit Function
    If Results.count = 0 Then IMAP_SmartDelPurge = True: Exit Function
    
    '   Delete all Items in the Collection
    '
    Dim oItem As Object
    For Each oItem In Results
        oItem.Delete
        DoEvents
    Next oItem

IMAP_SmartDelPurge = True
End Function

'   Purge All Known IMAP Folders
'
Public Function IMAP_SmartDelPurgeAll() As Boolean
IMAP_SmartDelPurgeAll = False

    Dim Paths() As String
    Paths = Folders_KnownPaths()
    
    Dim Index As Long
    For Index = 1 To UBound(Paths): Do
    
        Dim FolderPath As String
        FolderPath = Paths(Index)
        If Not IMAP_FolderPathIsIMAP(FolderPath) Then Exit Do   ' Next Index
        If Not IMAP_SmartDelPurge(FolderPath) Then Stop: Exit Function
        
    Loop While False: Next Index

IMAP_SmartDelPurgeAll = True
End Function
