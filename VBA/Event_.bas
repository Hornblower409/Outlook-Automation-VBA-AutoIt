Attribute VB_Name = "Event_"
Option Explicit
Option Private Module

' =====================================================================
'   Init
' =====================================================================

Public Function Event_InitApplication() As Boolean
Event_InitApplication = False

    '   Check the KnowPaths
    '
    If Not Folders_KnownPathsCheck Then Stop: Exit Function

    '   Init my Globals
    '
    If Not Globals_Init() Then Stop: Exit Function

    '   Backup the VbaProject.OTM file
    '   Backup the Master Cats collection
    '
    If Not File_BackupVBAProject_Trim Then Stop: Exit Function
    If Not Categories_MasterCatsBackup Then Stop: Exit Function

Event_InitApplication = True
End Function

' =====================================================================
'   Open Item Event
' =====================================================================

Public Function Event_OpenItemEvent(ByVal oItem As Object) As Boolean
Event_OpenItemEvent = False

    If CustForm_ClassIsCustom(oItem.MessageClass) Then
        If Not CustForm_OpenEvent(oItem) Then Exit Function
    Else
        If Not Event_OpenItem(oItem) Then Exit Function
    End If

Event_OpenItemEvent = True
End Function

' =====================================================================
'   Open Item
' =====================================================================

Private Function Event_OpenItem(ByVal oItem As Object) As Boolean
Event_OpenItem = False

    Dim InspectorRec As Inspector_InspectorRec
    If Not Inspector_RecFromItem(oItem, InspectorRec) Then Stop: Exit Function
    
    With InspectorRec
    
        '   Branch based on New or Existing Item
        '
        If oItem.EntryId = "" Then
            If Not Event_OpenItemNew(InspectorRec) Then Exit Function
        Else
            If Not Event_OpenItemExisting(InspectorRec) Then Exit Function
        End If
        
    End With

Event_OpenItem = True
End Function

Private Function Event_OpenItemNew(InspectorRec As Inspector_InspectorRec) As Boolean
Event_OpenItemNew = False

    With InspectorRec
    
        '   If RTF that can be converted to HTML - Do it.
        '   Add/Appy Hot Rod Normal Style
        '
        If Mail_IsRTF(.oItem) Then
            .oItem.BodyFormat = Outlook.olFormatHTML
        End If
        If Not Format_HotRodStyle(.oItem) Then Stop: Exit Function

    End With
    
Event_OpenItemNew = True
End Function

Private Function Event_OpenItemExisting(InspectorRec As Inspector_InspectorRec) As Boolean
Event_OpenItemExisting = False

    With InspectorRec

        '   Handle Special Cases
        '
        Select Case True
            Case (TypeOf .oItem Is Outlook.MailItem)
                If Not Event_OpenExistingMail(.oItem) Then Exit Function
            Case Else
                '   Continue
        End Select

    End With
    
Event_OpenItemExisting = True
End Function

' ---------------------------------------------------------------------
'   Existing MailItem Open
' ---------------------------------------------------------------------

Private Function Event_OpenExistingMail(ByVal MailItem As Outlook.MailItem) As Boolean
Event_OpenExistingMail = False

    '   If it is an UnSent Draft in the Inbox
    '
    If (MailItem.Parent.FolderPath = glbKnownPath_HomeInBox) Then
        If Not Mail_IsSent(MailItem) Then

            '   Move it back to Drafts and Open it from there
            '   Cancel this Open from the Inbox
            '
            If Not Draft_MoveToDrafts(MailItem) Then Stop: Exit Function
            Exit Function

        End If
    End If

Event_OpenExistingMail = True
End Function

' =====================================================================
'   Misc Item
' =====================================================================

' ---------------------------------------------------------------------
'   Send Item
' ---------------------------------------------------------------------

Public Function Event_SendItem(ByVal oItem As Object) As Boolean
Event_SendItem = False

    Dim Cancel As Boolean
    Send_ItemSend oItem, Cancel
    If Cancel Then Exit Function

Event_SendItem = True
End Function

' ---------------------------------------------------------------------
'   Add Item
' ---------------------------------------------------------------------

Public Function Event_AddItem(ByVal oItem As Object, ByVal oFolder As Outlook.Folder) As Boolean
Event_AddItem = False

    Select Case oFolder.FolderPath
        Case glbKnownPath_Projects
            If Not Event_AddItemProjects(oItem) Then Stop: Exit Function
        Case Else
            Stop: Exit Function
    End Select

Event_AddItem = True
End Function

Public Function Event_AddItemProjects(ByVal oItem As Object) As Boolean
Event_AddItemProjects = False

    Projects_ItemAdd oItem

Event_AddItemProjects = True
End Function

' ---------------------------------------------------------------------
'   Item BeforeDelete
' ---------------------------------------------------------------------

Public Function Event_BeforeDelete(ByVal MailItem As Outlook.MailItem) As Boolean
Const ThisProc = "Event_BeforeDelete"

    Event_BeforeDelete = True

    '   If the Item is in the Drafts folder - Move to Deleted Items
    '
    '       HACK - My Timer on Inspector Close looks for Drafts and moves them to the Inbox.
    '       But there is a delay on Delete before the Item is actually in Deleted Items
    '       and I was catching it before it made it. This just makes sure any deleted Item is not
    '       in Drafts when the Inspector closes so I won't catch it by mistake.
    '
    If MailItem.Parent.FolderPath = glbKnownPath_Drafts Then

        '   Move it to Delete Items
        '   Cancel the normal Delete
        '
        Dim MovedItem As Object
        Set MovedItem = MailItem.Move(Folders_KnownPath(glbKnownPath_Deleted))
        If MovedItem Is Nothing Then Stop: Exit Function
        Event_BeforeDelete = False

    End If

End Function

' ---------------------------------------------------------------------
'   Item PropertyChange
' ---------------------------------------------------------------------

Public Sub Event_PropChangePost(ByVal PostItem As Outlook.PostItem, ByVal IsStandardProp As Boolean, ByVal PropName As String)

    If Not CustForm_ClassIsCustom(PostItem.MessageClass) Then Exit Sub
    If Not CustForm_PropChange(PostItem, IsStandardProp, PropName) Then Stop: Exit Sub

End Sub

' ---------------------------------------------------------------------
'   Item CustomAction
' ---------------------------------------------------------------------

Public Sub Event_CustomActionPost(ByVal Action As Outlook.Action, ByVal Response As Outlook.PostItem, ByRef Cancel As Boolean)

    '   Get the Form from the Action
    '
    Dim oForm As Outlook.PostItem
    Set oForm = Action.Parent
    If oForm Is Nothing Then Stop: Exit Sub

    '   If Not a Custom Form - Done
    '
    If Not CustForm_ClassIsCustom(oForm.MessageClass) Then Exit Sub

    '   Ditch the Response
    '   Setup to Cancel the Action on return
    '
    If Not Response Is Nothing Then Response.Close Outlook.OlInspectorClose.olDiscard
    Cancel = True

    '   Form Specific Custom Action
    '
    If Not CustForm_CustAction(oForm, Action.Name) Then Stop: Exit Sub

End Sub

' ---------------------------------------------------------------------
'   Item Write
' ---------------------------------------------------------------------

Public Function Event_WriteItem(ByVal oItem As Object) As Boolean
Event_WriteItem = False

    '   Custom Form
    '
    If CustForm_ClassIsCustom(oItem.MessageClass) Then
        If Not CustForm_Write(oItem) Then Exit Function
    End If

Event_WriteItem = True
End Function

' =====================================================================
'   CmdButton Click
' =====================================================================

Public Sub Event_ClickCmdButton(ByVal oItem As Object, ByVal oCmdButton As MSForms.CommandButton)

    If CustForm_ClassIsCustom(oItem.MessageClass) Then
        If Not CustForm_ClickCmdButton(oItem, oCmdButton) Then Stop: Exit Sub
    End If

End Sub

' =====================================================================
'   Exit Event Scope - Inspector
' =====================================================================

Public Function Event_ExitEventScope(ByVal oItem As Object, ByVal AfterExit As Long) As Boolean
Const ThisProc = "Event_ExitEventScope"
Event_ExitEventScope = False

    '   Build my Timer
    '
    Dim Timer As Timer: Set Timer = New Timer
    With Timer

        .Name = ThisProc
        .Period = 25
        .Retries = 4
        Set .StashObj = oItem                                   ' My Item
        .StashVar = AfterExit                                   ' Where to go after I'm out of Event Scope
        .Callback AddressOf Event_ExitEventScopeTimer           ' My Timer Callback Proc

    End With

    '   Start my Timer
    '
    glbAppTimers.Add Timer
    Timer.Enable

Event_ExitEventScope = True
End Function

Private Sub Event_ExitEventScopeTimer(ByVal P1 As Long, ByVal P2 As Long, ByVal TimerId As Long, ByVal P3 As Long)
Const ThisProc = "Event_ExitEventScopeTimer"

    '   Disable the Timer while I'm running
    '
    Dim Timer As Timer: Set Timer = glbAppTimers.Timers(TimerId)
    If Timer Is Nothing Then Stop: GoTo Exit_Sub
    Timer.Disable

    '   Get my stuff from the Timer Stash
    '
    Dim oItem As Object
    Set oItem = Timer.StashObj
    If oItem Is Nothing Then Stop: GoTo Exit_Sub
    
    Dim AfterExit As Variant
    AfterExit = CLng(Timer.StashVar)

    '   Get the Item's InspectorRec
    '
    If Not Inspector_ItemInspectorExist(oItem) Then Stop: GoTo Exit_Sub
    Dim InspectorRec As Inspector_InspectorRec
    If Not Inspector_RecFromItem(oItem, InspectorRec) Then Stop: GoTo Exit_Sub
    If InspectorRec.oInspector Is Nothing Then Stop: GoTo Exit_Sub
    With InspectorRec
    
        '   Retry until the Item's Inspector Ribbon is enabled or Time Out
        '
        If Not Ribbon_Active(.oInspector) Then
    
            DoEvents
            Select Case Timer.Retry
                Case True
                    Timer.Enable
                Case Else
                    Msg_Box Proc:=ThisProc, Step:="Timer Retries", Text:= _
                        "Timer: " & Timer.Name & vbNewLine & _
                        "Inspector: " & .oInspector.Caption & glbBlankLine & _
                        "Timed Out waiting for the Inspector to enable the Ribbon."
                    glbAppTimers.Remove Timer
            End Select
            GoTo Exit_Sub
    
        End If
    
    End With

    '   Item's Inspector Ribbon is enabled - Go based on AfterExit
    '
    glbAppTimers.Remove Timer
    Select Case AfterExit
        Case glbExitEventScope_goNoWhere
            '   Continue
        Case glbExitEventScope_goCustFormOpen
            If Not CustForm_Open(oItem) Then GoTo Exit_Sub
        Case Else
            Stop: GoTo Exit_Sub
    End Select

Exit_Sub:

    glbAppTimers.Remove Timer

End Sub

' =====================================================================
'   Inspector Event Handlers
' =====================================================================

Public Function Event_CloseInspector(ByVal Item As Object, ByVal InspShadow As InspShadow) As Boolean
Event_CloseInspector = False

    '   If a Draft - Start the Timer to move it from Drafts to the Inbox
    '   once the Inspector is fully closed.
    '
    If TypeOf Item Is Outlook.MailItem Then
        If Draft_IsDraft(Item) Then
            If Not Draft_MoveToInbox(Item, InspShadow.ShadowKey) Then Stop: Exit Function
        End If
    End If

Event_CloseInspector = True
End Function

Public Function Event_ActivateInspector(ByVal Item As Object, ByVal InspShadow As InspShadow) As Boolean
Event_ActivateInspector = False

    '       SPOS - This should be done in a Response Open Events, but the HTML is not fully formed until
    '       after the Open has finished. So we have to check on every Activate.
    '
    If TypeOf Item Is Outlook.MailItem Then
        If glbCleanupAuto Then
            If Not InspShadow.CleanupResponseBypass Then
                If Not InspShadow.CleanupResponseDone Then
                    If Mail_IsResponse(Item) Then
                        Cleanup_Response Item
                        InspShadow.CleanupResponseDone = True
                    End If
                End If
            End If
        End If
    End If

Event_ActivateInspector = True
End Function

' =====================================================================
'   Explorer Event Helpers
' =====================================================================

'   Return an Explorer's first and only selected Item or Nothing
'
Public Function Event_SelectedExplorer(ByVal oExplorer As Outlook.Explorer, ByRef oSelected As Object) As Boolean
Event_SelectedExplorer = False

    '   Only want it if it is a single item from a standard Table View
    '
    Dim oSelection As Outlook.Selection
    Set oSelection = oExplorer.Selection
    If Not oSelection.count = 1 Then Exit Function
    If Not oSelection.Location = Outlook.olViewList Then Exit Function

    On Error Resume Next
        Set oSelected = oSelection.Item(1)
        Select Case Err.Number
            Case glbError_None
            Case Else
                Stop: Exit Function
        End Select
    On Error GoTo 0

Event_SelectedExplorer = True
End Function

'   Start a Timer to wait for an Explorer to be completley Activated
'   and the first selected Item (if any) to be Saved.
'
Public Function Event_ActivateExplorer(ByVal oExplShadow As ExplShadow, ByVal ExplShadowKey As String, ByVal oExplorer As Outlook.Explorer) As Boolean
Const ThisProc = "Event_ActivateExplorer"
Event_ActivateExplorer = False

    '   Build my Timer object
    '
    Dim Timer As Timer: Set Timer = New Timer
    With Timer

        .Name = ThisProc & "_" & ExplShadowKey
        .Period = 250
        .Retries = 4
        .Callback AddressOf Event_ActivateExplorer_Timer

        '   Stash a reference to the ExplShadow, ExplShadowKey, and Explorer for the Callback
        '
        .StashVar = Array(oExplShadow, ExplShadowKey, oExplorer)

    End With

    '   Start my Timer
    '
    glbAppTimers.Add Timer
    Timer.Enable

Event_ActivateExplorer = True
End Function

'   Timer Callback for Event_ActivateExplorer
'
Private Sub Event_ActivateExplorer_Timer(ByVal P1 As Long, ByVal P2 As Long, ByVal TimerId As Long, ByVal P3 As Long)
Const ThisProc = "Event_ActivateExplorer_Timer"

    '   Get my Timer and Disable it
    '
    Dim Timer As Timer: Set Timer = glbAppTimers.Timers(TimerId)
    If Timer Is Nothing Then Stop: Exit Sub
    Timer.Disable

    '   Get the info from the Timer Stash
    '
    Dim oExplShadow As ExplShadow:      Set oExplShadow = Timer.StashVar(0)
    Dim ExplShadowKey As String:        ExplShadowKey = CStr(Timer.StashVar(1))
    Dim oExplorer As Outlook.Explorer:  Set oExplorer = Timer.StashVar(2)

    '   If the Explorer is fully Activated ...
    '
    Dim ErrorMsg As String
    ErrorMsg = "Explorer to become the ActiveExplorer."
    If oExplorer Is Application.ActiveExplorer Then

        '   And If the Explorer's first selected Item (if any) is Saved - Done
        '
        ErrorMsg = "Explorer's first selected Item to be Saved."
        Dim oSelected As Object
        If Not Event_SelectedExplorer(oExplorer, oSelected) Then GoTo ExitSub
        If oSelected Is Nothing Then GoTo ExitSub
        If oSelected.Saved Then GoTo ExitSub

    End If

    '   Else
    '
    '       If any Retries left - Do it.
    '       Else Error Exit.
    '
    If Timer.Retry Then Timer.Enable: Exit Sub

    Msg_Box Proc:=ThisProc, Step:="Timer Retry", Text:= _
        "Timer: " & Timer.Name & vbNewLine & _
        "Explorer Title: " & oExplorer.Caption & vbNewLine & _
        "Explorer Folder: " & oExplorer.CurrentFolder.FolderPath & vbNewLine & _
        "Explorer View: " & oExplorer.CurrentView & vbNewLine & _
        "ExplShadowKey: " & ExplShadowKey & glbBlankLine & _
        "Timed Out waiting for the " & ErrorMsg

    GoTo ExitSub

ExitSub:

    glbAppTimers.Remove Timer
    oExplShadow.SelectionChange

End Sub

' =====================================================================
'   Delayed Explorer Activate
' =====================================================================

'   Give any current Explorer event time to finish
'   and then activate another Explorer.
'
Public Function Event_DelayedExplorerActivate(ByVal oExplorer As Outlook.Explorer) As Boolean
Const ThisProc = "Event_DelayedExplorerActivate"
Event_DelayedExplorerActivate = False

    '   Build my Timer object
    '
    Dim Timer As Timer: Set Timer = New Timer
    With Timer

        .Name = ThisProc
        .Period = 25
        .Retries = 0
        Set .StashObj = oExplorer
        .Callback AddressOf Event_DelayedExplorerActivate_Timer

    End With

    '   Start my Timer
    '
    glbAppTimers.Add Timer
    Timer.Enable

Event_DelayedExplorerActivate = True
End Function

Private Sub Event_DelayedExplorerActivate_Timer(ByVal P1 As Long, ByVal P2 As Long, ByVal TimerId As Long, ByVal P3 As Long)
Const ThisProc = "Event_DelayedExplorerActivate_Timer"

    '   Get my Timer and Disable it
    '
    Dim Timer As Timer: Set Timer = glbAppTimers.Timers(TimerId)
    If Timer Is Nothing Then Stop: Exit Sub
    Timer.Disable
    
    '   Get the Explorer to Activate from the Timer Stash
    '
    Dim oExplorer As Outlook.Explorer
    Set oExplorer = Timer.StashObj
    
    '   Kill my Timer
    '   Activate the Stashed Explorer
    '
    glbAppTimers.Remove Timer
    oExplorer.Activate

End Sub
