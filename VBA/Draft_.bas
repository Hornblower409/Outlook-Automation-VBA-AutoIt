Attribute VB_Name = "Draft_"
Option Explicit
Option Private Module

' ---------------------------------------------------------------------
'   Draft - Inbox to Drafts
'
'       Called by Mail Open. Moves a Saved but Unsent item from the Inbox
'       (user did a Save and Close on a new item, or I put it there previously)
'       back to Drafts. So Stupid will treat it like a real Draft item when he opens it.
'
' ---------------------------------------------------------------------
'
Public Function Draft_MoveToDrafts(ByVal Item As Outlook.MailItem) As Boolean
Const ThisProc = "Draft_MoveToDrafts"
Draft_MoveToDrafts = False
    
    '   Move to Drafts
    '
    Dim DraftItem As Outlook.MailItem
    Set DraftItem = Item.Move(Folders_KnownPath(glbKnownPath_Drafts))
    If DraftItem Is Nothing Then
        Msg_Box Proc:=ThisProc, Step:="Item.Move", Text:="Move from Inbox to Drafts failed"
        Stop: Exit Function
    End If
    
    '   Show it as a Draft item
    '
    If Inspector_ItemInspectorExist(DraftItem) Then Stop: Exit Function
    DraftItem.GetInspector.Activate
    
Draft_MoveToDrafts = True
End Function

' ---------------------------------------------------------------------
'   Draft - Drafts to Inbox
'
'       Called by Inspector Close. If it is an Unsent item in Drafts
'       move it to the Inbox. Requires a Timer Loop because I can't do
'       any of this from inside the Inspector/Item Close event.
'
' ---------------------------------------------------------------------
'
Public Function Draft_MoveToInbox(ByVal Item As Outlook.MailItem, ByVal InspShadowKey As String) As Boolean
Const ThisProc = "Draft_MoveToInbox"
Draft_MoveToInbox = False

    '   Build my Timer object
    '
    Dim Timer As Timer: Set Timer = New Timer
    With Timer
    
        .Name = ThisProc
        .Period = 250
        .Retries = 4
        .Callback AddressOf Draft_MoveToInbox_Timer
        
        '   Stash info on this Inspector (it's InspShadowKey) and the Item
        '   in a Variant Array in the Timer object. We can't hold on to the Item or the
        '   Inspector *references*, or bad things happen.
        '
        .StashVar = Array(InspShadowKey, Item.EntryId, Item.Parent.StoreID)
    
    End With
    
    '   Start my Timer
    '
    glbAppTimers.Add Timer
    Timer.Enable

Draft_MoveToInbox = True
End Function
Private Sub Draft_MoveToInbox_Timer(ByVal P1 As Long, ByVal P2 As Long, ByVal TimerId As Long, ByVal P3 As Long)
Const ThisProc = "Draft_MoveToInbox_Timer"

    '   Get my Timer and Disable it
    '
    Dim Timer As Timer: Set Timer = glbAppTimers.Timers(TimerId)
    If Timer Is Nothing Then Stop: Exit Sub
    Timer.Disable
   
    '   Get the InspShadowKey and the Item info from the Timer Stash
    '
    Dim InspShadowKey As String:    InspShadowKey = CStr(Timer.StashVar(0))
    Dim ItemEntryId As String:      ItemEntryId = CStr(Timer.StashVar(1))
    Dim FolderStoreID As String:    FolderStoreID = CStr(Timer.StashVar(2))
    
    '   If the Inspector is closed
    '
    If glbAppShadows.InspShadows(InspShadowKey) Is Nothing Then
    
        '   Get the Item from the Stashed info
        '
        Dim Item As Outlook.MailItem
        Set Item = Misc_GetItemFromID(ItemEntryId, FolderStoreID)
        
        '   If still a real Draft - move it to the Inbox
        '
        Dim InboxItem As Object
        If Draft_IsDraft(Item) Then
            Set InboxItem = Item.Move(Folders_KnownPath(glbKnownPath_HomeInBox))
            If InboxItem Is Nothing Then Stop: Exit Sub
        End If
        
        '   Close up and go home
        '
        glbAppTimers.Remove Timer
        Exit Sub
    
    End If

    '   The Inspector has not finished closing (i.e. the Inspector
    '   Shadow still exist) - Restart the Timer Loop or Timeout
    '
    If Timer.Retry Then Timer.Enable: Exit Sub
    
    glbAppTimers.Remove Timer
    Msg_Box Proc:=ThisProc, Step:="Retries", Text:= _
        "Timer: " & Timer.Name & vbNewLine & _
        "ItemEntryID: " & ItemEntryId & vbNewLine & _
        "InspShadowKey: " & InspShadowKey & glbBlankLine & _
        "Timed Out waiting for the Inspector to close."
    
End Sub

'   Is the Item a real Draft?
'
Public Function Draft_IsDraft(ByVal Item As Object) As Boolean
Draft_IsDraft = False

    '   Make sure we got a valid Mail
    '
    If Item Is Nothing Then Exit Function
    If Not (TypeOf Item Is Outlook.MailItem) Then Exit Function
    
    '   If Item has been moved, deleted, or sent - done
    '
    On Error Resume Next
            
        If Item.Parent Is Nothing Then Exit Function
        If Err.Number = glbError_ItemIsZombie Then Exit Function
        If Err.Number <> glbError_None Then Stop: Exit Function
        
    On Error GoTo 0
        
    '   If not in the Drafts folder, not Saved or Sent - done
    '
    If Not (Item.Parent.FolderPath = glbKnownPath_Drafts) Then Exit Function
    If Not Item.Saved Then Exit Function
    If Mail_IsSent(Item) Then Exit Function

Draft_IsDraft = True
End Function
