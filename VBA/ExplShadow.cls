VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExplShadow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents MyExplorer       As Outlook.Explorer
Attribute MyExplorer.VB_VarHelpID = -1

Private WithEvents MailItem         As Outlook.MailItem
Attribute MailItem.VB_VarHelpID = -1
Private WithEvents MeetingItem      As Outlook.MeetingItem
Attribute MeetingItem.VB_VarHelpID = -1
Private WithEvents PostItem         As Outlook.PostItem
Attribute PostItem.VB_VarHelpID = -1

Private MyExplShadowKey             As String
'

' =====================================================================
'               Explorer Events
' =====================================================================

Public Function Initialize(ByVal Explorer As Explorer, ByVal ExplShadowKey As Long) As Boolean
Initialize = False

    '   Hook MyExplorer Events
    '
    Set MyExplorer = Explorer

    '   Update the current Explorer Selection.
    '   And hook the current Item events.
    '
    MyExplorer_SelectionChange

    '   Housekeeping
    '
    MyExplShadowKey = CStr(ExplShadowKey)
    glbAppShadows.ExplShadowsAdd Me, MyExplShadowKey
    
Initialize = True
End Function

'   Explorer Close
'
Private Sub MyExplorer_Close()

    '   Make myself disappear
    '
    glbAppShadows.ExplShadowsRemove MyExplShadowKey

End Sub

'   Explorer Before Folder Switch
'
'       2023-02-03 - Catch clicking on a date in the To Do Bar
'       (when viewing the Inbox) which causes the Inbox folder to change to the Calander.
'
Private Sub MyExplorer_BeforeFolderSwitch(ByVal NewFolder As Object, Cancel As Boolean)
Const ThisProc = "ExplShadow.MyExplorer_BeforeFolderSwitch"

    '   If MyExplorer or MyExplorer.CurrentFolder not defined - let it change.
    '
    If MyExplorer Is Nothing Then Exit Sub
    If MyExplorer.CurrentFolder Is Nothing Then Exit Sub
    
    '   If the Explorer Current Folder is NOT the Home Inbox - let it change.
    '
    If Not (MyExplorer.CurrentFolder.FolderPath = glbKnownPath_HomeInBox) Then Exit Sub
    
    '   If the new Folder is NOT the Calendar - let it change.
    '
    If Not (NewFolder.Name = "Calendar") Then Exit Sub
    
    '   Set to Cancel the Inbox Explorer Folder Change on exit
    '
    Cancel = True
    
    '   Walk the open Explorers looking for a Calendar Explorer
    '
    Dim CollectionExplorer As Outlook.Explorer
    Dim CalendarExplorer As Outlook.Explorer
    
    For Each CollectionExplorer In Application.Explorers
        
        If CollectionExplorer.CurrentFolder.Name = "Calendar" Then
            Set CalendarExplorer = CollectionExplorer
            Exit For
        End If
        
    Next
       
    '   If there is NOT a Calendar Explorer - error exit
    '
    If CalendarExplorer Is Nothing Then
        Msg_Box Proc:=ThisProc, Step:="Find Calendar Explorer", Text:="Can not switch to a Calendar Explorer because there is not one open."
        Exit Sub
    End If
    
    '   Activate the Calendar Explorer
    '
    '   SPOS 2025-10-01 - When the Navigation Bar is set to Shortcuts (v Folders) and minimized
    '   on return from here it jumps back to the Inbox and shows the Shortcuts pop-up.
    '
    '   From calling an Activate from inside or another event handler. Switched
    '   to a Timer to give the BeforeFolderSwitch event time to finish first.
    '
    '   It still flashes the Nav Bar shortcuts but so fast you can't really see it.
    '
    If Not Event_DelayedExplorerActivate(CalendarExplorer) Then Stop: Exit Sub
    
End Sub

' =====================================================================
'   Explorer Selection Change
'
'       So I can catch a Reply, Reply All, or Forward on a single selected
'       Post or Mail Item in an Explorer and run Cleanup on the Response.
'
'       2025-02-01 - SPOS - Dirty Zombie
'
'       If the Item was modified and then closed with Save Changes = No it is causing a "Dirty Zombie".
'       i.e. The actual Item on disk (and as shown in the Explorer) reverted to the previous Saved version.
'       But if you Open the Item again, you may get the dirty version.
'
'       Moving the Explorer selection to a different item (and lots of other stuff) will (sometimes) cause
'       him to discard the Dirty Zombie and the next time you open the Item you'll see the saved version.
'
'       From my testing it looks like he is caching the Explorer Selected Item. The only way (I could find)
'       to make him let go is to get thru a COMPLETE Explorer_Activate without holding any references
'       to the current Explorer selection.
'
'       So - In the Explorer_Activate event, if my reference is to a Dirty Zombie, I release my reference
'       and start a Timer to callback after the Activate has completley finished (and the cache
'       has been updated?). Then I put my hooks in place again.
'
'       See: Card Software | Outlook | VBA - Dirty Zombie.
'
' =====================================================================

'   Explorer Activate Event
'
Private Sub MyExplorer_Activate()
Const ThisProc = "MyExplorer_Activate"

    '   myItem = Currently hooked Mail. Meeting, or Post
    '
    Dim MyItem As Object
    If Not MailItem Is Nothing Then Set MyItem = MailItem
    If Not MeetingItem Is Nothing Then Set MyItem = MeetingItem
    If Not PostItem Is Nothing Then Set MyItem = PostItem

    '   If no current Hooks - Done
    '
    If MyItem Is Nothing Then Exit Sub

    '   If myItem is Saved or Deleted - Done
    '
    On Error Resume Next

        If MyItem.Saved Then Exit Sub
        Select Case Err.Number
            Case glbError_None
            Case glbError_ItemIsZombie
                Exit Sub
            Case Else
                Stop: Exit Sub
        End Select

    On Error GoTo 0

    '   If my Item is open in any Inspector - Done
    '
    '       It could be Dirty because it actually WAS changed in an Imspector
    '       that hasn't closed yet.
    '
    If Inspector_ItemInspectorExist(MyItem) Then Exit Sub

    '   We Got One! (A Dirty Zombie).
    '
    ' Debug.Print ThisProc & ". " & MyExplShadowKey & ". We Got One! " & myItem.Subject

    '   Release all my References and Start a Timer that returns at
    '   ExplShadow_ActivateCallback once the Explorer is fully Activated and the
    '   Item is no longer Dirty.
    '
    Set MailItem = Nothing
    Set MeetingItem = Nothing
    Set PostItem = Nothing

    If Not Event_ActivateExplorer(Me, MyExplShadowKey, MyExplorer) Then Stop: Exit Sub

End Sub

'   Explorer SelectionChange Method and Event
'
'   If thre is only one Item selected and it's a Mail, Meeting, or Post - Hook It.
'
Public Sub SelectionChange():   MyExplorer_SelectionChange:    End Sub
Private Sub MyExplorer_SelectionChange()

    '   Clear any hooks.
    '
    Set MailItem = Nothing
    Set MeetingItem = Nothing
    Set PostItem = Nothing

    '   Get myExplorer's selection
    '
    Dim oSelected As Object
    If Not Event_SelectedExplorer(MyExplorer, oSelected) Then GoTo ExitSub
    If oSelected Is Nothing Then GoTo ExitSub

    '   Set myExplorer Event hooks.
    '   (Ignore any 430 errors - Item is defined but incomplete)
    '
    On Error Resume Next

        If TypeOf oSelected Is MailItem Then Set MailItem = oSelected
        If TypeOf oSelected Is MeetingItem Then Set MeetingItem = oSelected
        If TypeOf oSelected Is PostItem Then Set PostItem = oSelected
        
        Select Case Err.Number
            Case glbError_None
            Case glbError_Automation430
                DoEvents
            Case Else
                Stop: Exit Sub
        End Select

    On Error GoTo 0

ExitSub:

End Sub

' =====================================================================
'   Item Events
' =====================================================================

'   Response
'
Private Sub MailItem_Reply(ByVal Response As Object, Cancel As Boolean):        Response_Main MailItem, Response, Cancel, glbResponse_Reply:        End Sub
Private Sub MeetingItem_Reply(ByVal Response As Object, Cancel As Boolean):     Response_Main MeetingItem, Response, Cancel, glbResponse_Reply:     End Sub
Private Sub PostItem_Reply(ByVal Response As Object, Cancel As Boolean):        Response_Main PostItem, Response, Cancel, glbResponse_Reply:        End Sub
Private Sub MailItem_ReplyAll(ByVal Response As Object, Cancel As Boolean):     Response_Main MailItem, Response, Cancel, glbResponse_ReplyAll:     End Sub
Private Sub MeetingItem_ReplyAll(ByVal Response As Object, Cancel As Boolean):  Response_Main MeetingItem, Response, Cancel, glbResponse_ReplyAll:  End Sub
Private Sub PostItem_ReplyAll(ByVal Response As Object, Cancel As Boolean):     Response_Main PostItem, Response, Cancel, glbResponse_ReplyAll:     End Sub
Private Sub MailItem_Forward(ByVal Response As Object, Cancel As Boolean):      Response_Main MailItem, Response, Cancel, glbResponse_Forward:      End Sub
Private Sub MeetingItem_Forward(ByVal Response As Object, Cancel As Boolean):   Response_Main MeetingItem, Response, Cancel, glbResponse_Forward:   End Sub
Private Sub PostItem_Forward(ByVal Response As Object, Cancel As Boolean):      Response_Main PostItem, Response, Cancel, glbResponse_Forward:      End Sub

