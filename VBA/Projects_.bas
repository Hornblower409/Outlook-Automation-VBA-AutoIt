Attribute VB_Name = "Projects_"
Option Explicit
Option Private Module

' =====================================================================
'   Cat and Move
' =====================================================================

' ---------------------------------------------------------------------
'   Cat And Move To JUNK
'   Replaces any existing Cats with just JUNK
' ---------------------------------------------------------------------
Public Function Projects_CatAndMoveJunk() As Boolean
Const ThisProc = "Projects_CatAndMoveJunk"
Projects_CatAndMoveJunk = False

    Dim Cats As String
    Cats = glbCatJunk

    '   If an Explorer - Can't be open to Projects
    '
    If (TypeOf ActiveWindow Is Outlook.Explorer) And (ActiveExplorer.CurrentFolder.FolderPath = glbKnownPath_Projects) Then
        Msg_Box Proc:=ThisProc, Text:="You want to do a Cat And Move to Projects? You are already IN Projects dufus."
        Exit Function
    End If
        
    '   Assign Junk to selected items
    '
    If Not Categories_AssignFixed(Cats) Then Exit Function
    
    '   Do the Move
    '
    Projects_CatAndMove_Selected FollowUp:=False
    
Projects_CatAndMoveJunk = True
End Function

' ---------------------------------------------------------------------
'   Cat and Move (Show Cat Dialog for selection)
' ---------------------------------------------------------------------
'
Public Function Projects_CatAndMove(ByVal FollowUp As Boolean) As Boolean
Const ThisProc = "Projects_CatAndMove"
Projects_CatAndMove = False

    '   If FollowUp - Can we do it?
    '
    If FollowUp Then If Not FollowUp_Precheck(NoIMAP:=False, Caller:=ThisProc) Then Exit Function

    '   If an Explorer - Can't be open to Projects
    '
    If (TypeOf ActiveWindow Is Outlook.Explorer) And (ActiveExplorer.CurrentFolder.FolderPath = glbKnownPath_Projects) Then
        Msg_Box Proc:=ThisProc, Text:="You want to do a Cat And Move to Projects? You are already IN Projects dufus."
        Exit Function
    End If
        
    '   Get Cats for selected Item(s)
    '
    If Not Categories_Assign Then Exit Function
    
    '   Do the Move
    '
    If Not Projects_CatAndMove_Selected(FollowUp) Then Exit Function
    
Projects_CatAndMove = True
End Function

' ---------------------------------------------------------------------
'   Move selected items to the Projects folder
' ---------------------------------------------------------------------
Private Function Projects_CatAndMove_Selected(ByVal FollowUp As Boolean) As Boolean
Const ThisProc = "Projects_CatAndMove_Selected"
Projects_CatAndMove_Selected = False

    '   Inspector
    '
    If TypeOf ActiveWindow Is Outlook.Inspector Then
        
        If Not Projects_CatAndMove_Item(ActiveInspector.CurrentItem, FollowUp) Then Exit Function
        Projects_CatAndMove_Selected = True
        Exit Function
        
    End If
    
    '   Explorer
    '
    If TypeOf ActiveWindow Is Outlook.Explorer Then
    
        '   Get the Explorer Selection
        '
        Dim ExpSelection As Outlook.Selection
        Set ExpSelection = ActiveExplorer.Selection
        
        '   Must have something selected
        '
        If ExpSelection.count = 0 Then
            Msg_Box Proc:=ThisProc, Text:="Cat and Move. No items selected."
            Exit Function
        End If
        
        '   Walk the selection and do the Move
        '
        Dim ItemsMoved As Integer
        ItemsMoved = 0
        
        Dim SelItem As Object
        For Each SelItem In ExpSelection
            
            If Not Projects_CatAndMove_Item(SelItem, FollowUp) Then
                Exit For
            Else
                ItemsMoved = ItemsMoved + 1
            End If
            
        Next SelItem
        
        '   Check all moved
        '
        If ItemsMoved <> ExpSelection.count Then
            Msg_Box Proc:=ThisProc, Text:="Cat and Move. Not all items moved. Only " & ItemsMoved & " items moved to Projects."
            Exit Function
        End If
        
        '   Success
        '
        Projects_CatAndMove_Selected = True
        Exit Function
            
    End If
    
End Function

' ---------------------------------------------------------------------
'   Move an Item to the Projects folder
' ---------------------------------------------------------------------
'
Private Function Projects_CatAndMove_Item(ByVal SrcItem As Object, ByVal FollowUp As Boolean) As Boolean
Const ThisProc = "Projects_CatAndMove_Item"
Projects_CatAndMove_Item = False

    '   If {Item Type/Class} - Can not be moved to Projects
    '
    Select Case True
    
        '   Appointment or Meeting - No can do
        '
        Case (TypeOf SrcItem Is Outlook.AppointmentItem)
            Msg_Box Proc:=ThisProc, _
                Text:="Item is an Appointment or Meeting. Can not be moved to Projects." & glbBlankLine & _
                "(To Archive an Appointment or Meeting - Move it to the 'Calendar - Archive' folder).", _
                Subject:=SrcItem.Subject
            Exit Function
        
        Case Else
            '   Continue
            
    End Select
    
    '   If a Zombie - Ignore it
    '
    If Misc_ItemIsZombie(SrcItem) Then Exit Function
    
    '   Must have at least one Category assigned
    '
    If Not Projects_HasCats(SrcItem) Then Exit Function

    '   Must not already be in the Projects Folder
    '
    Dim SrcFolder As Outlook.Folder
    If Not Folders_Item(SrcItem, SrcFolder) Then Stop: Exit Function
    If SrcFolder.FolderPath = glbKnownPath_Projects Then
        Msg_Box Proc:=ThisProc, Text:="Cat and Move. Item already in Projects.", Subject:=SrcItem.Subject
        Exit Function
    End If
    
    '   If {Item Type/Class} - Must have been Sent
    '
    Select Case True
    
        '   Post - Let it go.
        '
        Case (SrcItem.Class = Outlook.olPost)
            '   Continue
        
        '   Else - Must have been sent
        '
        Case Else
            If Not Mail_IsSent(SrcItem) Then
                Msg_Box Proc:=ThisProc, Text:="Mail Item has not been Sent. Can not be moved to Projects.", Subject:=SrcItem.Subject
                Exit Function
            End If
        
    End Select
    
    '   If it will have a FollowUp in Projects - Get and stash the FollowUp param
    '
    If FollowUp Then If Not FollowUp_SetStash(SrcItem) Then Exit Function
    
    '   Move It with Trap and Check
    '
    Dim ProjectsFolder As Outlook.Folder
    Set ProjectsFolder = Folders_KnownPath(glbKnownPath_Projects)
    Dim ProjectItem As Object
    
    ' ---------------------------------------------------------------------
    '   2025-05-06 - Replaced with Copy and SmartDel/Delete
    '
    '    On Error GoTo CatAndMoveError
    '        Set ProjectItem = SrcItem.Move(ProjectsFolder)
    '    On Error GoTo 0
    '    If ProjectItem Is Nothing Then Stop: Exit Function
    ' ---------------------------------------------------------------------
    '
    Dim FileSpec As String
    If Not File_SaveToTemp(SrcItem, FileSpec) Then Stop: Exit Function
    If Not File_LoadFromFile(ProjectItem, FileSpec, Folders_KnownPath(glbKnownPath_Projects)) Then Stop: Exit Function
    
    '   2025-07-20 - Added On Error wrapper
    '
    On Error GoTo CatAndMoveError
        If Not IMAP_SmartDelCat(SrcItem) Then SrcItem.Delete
    On Error GoTo 0
    
    '   Normal Exit
    '
    Projects_CatAndMove_Item = True
    Exit Function
    
CatAndMoveError:

    If Not Projects_CatAndMoveError(SrcItem, SrcFolder, Err) Then Exit Function
    Projects_CatAndMove_Item = True
    
End Function

'   Handle CatAndMove Move Errors:
'
'       glbError_CopiedNotMoved - Item was copied not moved. Can't delete original. (aka Dana eMail Error)
'       glbError_ItemIsZombie - Item has already been deleted but Inspector is still Open.
'
Private Function Projects_CatAndMoveError(ByVal SrcItem As Object, ByVal SrcFolder As Outlook.Folder, ByVal Error As VBA.ErrObject) As Boolean
Const ThisProc = "Projects_CatAndMoveError"
Projects_CatAndMoveError = False

    Select Case Error.Number
    
        '   SPOS - Item has been copied to Projects but Stupid can't delete the Original.
        '   Only happens on an IMAP Folder and only on certain emails (e.g. From Dana)
        '
        '   The Folder/Explorer will still have an Original with the same MsgId.
        '   Any Inspector has become a Zombie (Item has has been moved or deleted).
        '
        Case glbError_CopiedNotMoved, glbError_IMAP_NoUIDDelete, glbError_ItemIsZombie
        
            Select Case Msg_Box( _
                Proc:=ThisProc, Step:="Dana Error (Can't delete original)", _
                Icon:=vbQuestion, Buttons:=vbYesNoCancel, Default:=vbDefaultButton1, _
                Subject:=SrcItem.Subject, _
                Text:="Cleanup any leftovers? (Cancel = Stop)")
            Case vbYes
                ' Continue
            Case vbNo
                Exit Function
            Case vbCancel
                Stop: Exit Function
            End Select

            '   Delete anything in the SrcFolder with this MsgId
            '
            Dim MsgId As String
            MsgId = IMAP_MsgID(SrcItem)
            If MsgId = "" Then Stop: Exit Function
            If Not IMAP_DeleteByMsgId(MsgId, SrcFolder) Then Stop: Exit Function
            
            '   If the ActiveInspector is this Item - Try a Close with Discard
            '
            '       Should check for this Item open in ANY Inspector. But there
            '       are limits to how far to chase a bug that happens so rarely.
            
            If TypeOf ActiveWindow Is Outlook.Inspector Then
                If (ActiveInspector.CurrentItem Is SrcItem) Then
                    On Error Resume Next
                    ActiveInspector.Close Outlook.OlInspectorClose.olDiscard
                    On Error GoTo 0
                End If
            End If
            
'        ' 2025-08-10 - Zombie goes thru full processing above. Dana's were comming straight here.
'        '
'        '   It's a Zombie (Item has has been moved or deleted).
'        '
'        Case glbError_ItemIsZombie
'
'            '   If an Inspector - Try a Close with Discard
'            '   If an Explorer - Can't happen?
'            '
'            If TypeOf ActiveWindow Is Outlook.Inspector Then
'                On Error Resume Next
'                ActiveInspector.Close Outlook.OlInspectorClose.olDiscard
'                On Error GoTo 0
'            Else
'                Stop: Exit Function
'            End If
            
        Case Else
            Stop: Exit Function
            
    End Select

Projects_CatAndMoveError = True
End Function

'---------------------------------------------------------------------
'   Cat And Send the Current Inspector
'
'   Cats will get stashed in BillingInformation and then stripped by Send_CatAndSend (on any Send Event).
'   Copy of the Sent Item will go to Projects. When the Sent Item shows up in Project, Projects_ItemAdd
'   event trigger will put the Cats back on.
'
' ---------------------------------------------------------------------
'
Public Function Projects_CatAndSend(ByVal FollowUp As Boolean) As Boolean
Const ThisProc = "Projects_CatAndSend"

    '   If FollowUp - Can we do it?
    '
    If FollowUp Then If Not FollowUp_Precheck(NoIMAP:=False, Caller:=ThisProc) Then Exit Function
    
    Dim InspItem As Object
    Dim InspType As String
    
    '   Active Window must be an Inspector
    '
    If Not Misc_GetActiveItem(InspItem, InspectorOnly:=True) Then Exit Function
    
    '   Normal items must not be sent.
    '   Appointment Item must be Meeting (have attendees) and have the special "Invitations Sent" property False.
    '
    InspType = TypeName(InspItem)
    Select Case InspType
    
        Case Is = "AppointmentItem"
        
            If InspItem.MeetingStatus = Outlook.olNonMeeting Then
                Msg_Box Text:="Appointment items (no attendees) can not be sent", Proc:=ThisProc
                Exit Function
            End If
        
            If Misc_CalendarInviteSent(InspItem) Then
                Msg_Box Text:="Meeting Invites have already been sent for this Appointment", Proc:=ThisProc
                Exit Function
            End If
        
        Case Is = "MailItem", "TaskItem", "MeetingItem", "PostItem"
        
            If Mail_IsSent(InspItem) = True Then
                Msg_Box Text:="Inspector Item has already been sent", Proc:=ThisProc
                Exit Function
            End If
            
        Case Else
        
            Msg_Box Text:="Inspector Item is an invalid type for Cat And Send. Type : " & TypeName(InspItem), Proc:=ThisProc
            Exit Function
            
    End Select

    '   Set Projects as where to save Sent
    '
    Dim ProjectsFolder As Outlook.Folder
    Set ProjectsFolder = Folders_KnownPath(glbKnownPath_Projects)
    
    On Error Resume Next
    
        Set InspItem.SaveSentMessageFolder = ProjectsFolder
        Select Case Err.Number
            Case glbError_None
            Case glbError_PropertyNotFound
                Msg_Box Proc:=ThisProc, Step:="Set SaveSentMessageFolder", _
                        Text:="Item does not support the SaveSentMessageFolder property. Can not be used in a Cat And Send."
                Exit Function
            Case Else
                Stop: Exit Function
        End Select
        
    On Error GoTo 0
                
    '   Get Cats
    '
    '       Do not stash cats in BI at this time because Send_CatAndSend
    '       (triggered on a Send Event) needs them in place. He will stash in BI after
    '       he has done his thing.
    '
    If Not Categories_Assign() Then Exit Function
    If InspItem.Categories = "" Then Exit Function
    
    '   Get FollowUp and stash in BI
    '
    If FollowUp Then If Not FollowUp_SetStash(InspItem) Then Exit Function
                
    ' InspItem.Send will bypass Spell Check before send hence the use of ExecuteMSO
    '
    If Not Ribbon_ExecuteMSO(Application.ActiveInspector, glbidMSO_SendDefault) Then Stop: Exit Function

End Function

' ---------------------------------------------------------------------
'   Projects Item Add
'
'       SPOS - This is really AFTER ItemAdd. There is no BEFORE ItemAdd. So there is no
'       Cancel return like on a Send event. I can not Stop them from adding an Item.
'
' ---------------------------------------------------------------------
'
Public Sub Projects_ItemAdd(ByVal Item As Object)
Const ThisProc = "Projects_ItemAdd"

    '   Verify Item is still valid (2025-02-10)
    '
    On Error Resume Next
        If Item.OutlookVersion = "<The message you specified cannot be found.>" Then Exit Sub
        If Err.Number <> glbError_None Then Exit Sub
    On Error GoTo 0
    
    '   Recover any Cats or FollowUp from BillingInfo
    '
    If Not Projects_ItemAddBI(Item) Then Exit Sub
    
    '   If it has Cats other than just FollowUp - Done
    '
    If (Item.Categories <> "") And (Item.Categories <> glbCatFollowUp) Then Exit Sub
    
    '   If Item has no Cats or just Follow Up - Open an Inspector on the Item and ask for Cats
    '
    Item.GetInspector.Activate
    Do
    
        If (Item.Categories <> "") And (Item.Categories <> glbCatFollowUp) Then Exit Do
        Select Case Msg_Box( _
                Proc:=ThisProc, Step:="Cats Check", _
                Subject:=Item.Subject, _
                Icon:=vbQuestion, Buttons:=vbYesNo, Default:=vbDefaultButton1, _
                Text:="Item added to Projects has no Categories." & glbBlankLine & _
                      "Assign Categories now?" & vbNewLine & _
                      "(Answering No will assign it the Cat '" & glbCatNoCats & "'.)" _
                )
            Case vbYes
                Categories_ShowCategoriesDialog Item
            Case vbNo
                Categories_AddCat Item, glbCatNoCats
                Item.Save
        End Select
        
    Loop
    
    Item.Close Outlook.olSave
    
End Sub

'   Process info stashed in the Billing Information
'
Private Function Projects_ItemAddBI(ByVal Item As Object) As Boolean
Projects_ItemAddBI = False

    '   If nothing in BI - Done
    '
    If Item.BillingInformation = "" Then
        Projects_ItemAddBI = True
        Exit Function
    End If

    ' ---------------------------------------------------------------------
    '   Cats
    ' ---------------------------------------------------------------------

    '   Get any Cats in BI
    '
    Dim CatsString As String
    If Not Mail_BIGet(Item, glbBIInx_Cats, CatsString) Then Stop: Exit Function
    
    '   Split the BI Cats string
    '   Add the Cats to the Item
    '
    Dim Cats As Variant
    Cats = Split(CatsString, glbCatSep)
    Dim Cat As Variant
    For Each Cat In Cats
        Categories_AddCat Item, CStr(Cat)
    Next Cat
    
    ' ---------------------------------------------------------------------
    '   Follow Up
    ' ---------------------------------------------------------------------
    
    '   Get the Follow Up param
    '
    Dim FlagRequest As String
    Dim ReminderTime As String
    If Not Mail_BIGet(Item, glbBIInx_FlagRequest, FlagRequest) Then Stop: Exit Function
    If Not Mail_BIGet(Item, glbBIInx_ReminderTime, ReminderTime) Then Stop: Exit Function
    
    '   If Follow Up defined
    '
    '       Set the Item Follow Up
    '       Add a Follow Up Cat
    '
    If FlagRequest <> "" Then
        If Not FollowUp_SetSave(Item, FlagRequest, ReminderTime) Then Stop: Exit Function
        Categories_AddCat Item, glbCatFollowUp
    End If
    
    ' ---------------------------------------------------------------------
    '   Done and Save
    ' ---------------------------------------------------------------------
    
    Item.BillingInformation = ""
    Item.Save

Projects_ItemAddBI = True
End Function

' ---------------------------------------------------------------------
'   Projects - Cat Check
'
'   Check for Project Items with no Cats.
'   Called from Scheduled Task Outlook_ProjectsCatCheck
'
' ---------------------------------------------------------------------
Public Function Projects_CatCheck() As Boolean
Projects_CatCheck = False

    '   Find any items with No Cats
    '
    Dim SQLRestrict As String
    SQLRestrict = "@SQL=" & glbQuote & glbPropTag_Categories & glbQuote & " IS NULL"
    Dim Results As VBA.Collection
    If Not Collection_FromRestrict(SQLRestrict, Folders_KnownPath(glbKnownPath_Projects), Results) Then Stop: Exit Function
    
    '   If none - Done
    '
    If Results.count = 0 Then
        Projects_CatCheck = True
        Exit Function
    End If
    
    '   For any No Cats - Hand off to the normal ItemAdd event handler
    '
    Dim oItem As Object
    For Each oItem In Results
        Projects_ItemAdd oItem
    Next oItem

Projects_CatCheck = True
End Function

' =====================================================================
'   Misc Projects Routines
' =====================================================================

' ---------------------------------------------------------------------
' Check if a message has categories
'
'   Show warning if no categories and allow add now.
'   Return TRUE if has categories.
'   Return FALSE if no categories or they picked Cancel.
' ---------------------------------------------------------------------
'
Public Function Projects_HasCats(ByVal Item As Object) As Boolean
Const ThisProc = "Projects_HasCats"

Repeat:

    ' Default return to TRUE (OK)
    Projects_HasCats = True
        
    ' If already has Cats - we're done
    If Item.Categories <> "" Then Exit Function
    
    ' Warn about no CATs and let them decide
    Select Case Msg_Box(Text:="Item must have at least one Category assigned." & glbBlankLine & "Assign Categories now?", Subject:=Item.Subject, Icon:=vbQuestion, Buttons:=vbOKCancel, Default:=vbDefaultButton1, Proc:=ThisProc)
        Case vbCancel
            Projects_HasCats = False
            Exit Function
        Case vbOK
            Categories_ShowCategoriesDialog Item
            GoTo Repeat
    End Select
    
End Function

