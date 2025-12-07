Attribute VB_Name = "FollowUp_"
Option Explicit
Option Private Module

' ---------------------------------------------------------------------
'   Testing on Flag and Follow Up
'
'   Set all the folloing:
'
'       FlagIcon = Outlook.olRedFlagIcon
'       FlagRequest = "{My Text}"  (Called "Follow Up Flag" in the GUI)
'       FlagStatus = Outlook.olFlagMarked
'       IsMarkedAsTask = True
'       ReminderSet = True
'       ReminderTime = {My Date and Time}
'       TaskCompletedDate = Nothing #4501-01-01#
'       TaskDueDate = Nothing #4501-01-01#
'       TaskStartDate = Nothing #4501-01-01#
'       TaskSubject = Item.Subject
'
'       Makes it "Starred" in GMail. Not sure which value he is looking at.
'
'   Move with Ctrl-Shift-V, or CaM to Projects and it's all gone except for:
'
'       FlagStatus
'       IsMarkedAsTask
'       TaskSubject
'
'   Set on a new Item gets you a -2147467263 (&H80004001)
'   "Draft items cannot be marked. MarkAsTask is only valid on items that have been sent or received."
'
' ---------------------------------------------------------------------

' ---------------------------------------------------------------------
'   Public Entry Points
' ---------------------------------------------------------------------
'
Public Function FollowUp_Set(Optional ByVal Caller As String = "FollowUp_Set") As Boolean
Const ThisProc = "FollowUp_Set"
FollowUp_Set = False

    '   Get the Active Object
    '
    Dim Item As Object
    If Not Misc_GetActiveItem(Item, Caller) Then Exit Function
    
    '   Can't Set on a Draft or Outbox (being composed) Item
    '
    '   Not part of the PreCheck beacuse CatAndSend doesn't have this restriction
    '   and he calls PreCheck.
    '
    Dim Folder As Outlook.Folder
    If Not Folders_Item(Item, Folder) Then Stop: Exit Function
    If (Folder.FolderPath = glbKnownPath_Drafts) Or (Folder.FolderPath = glbKnownPath_Outbox) Then
        Msg_Box _
            Proc:=Caller, _
            Step:="Drafts & Outbox Check", _
            Text:="Can not Set a Follow Up with Reminder for an Item in the Drafts Folder or one being composed in the Outbox folder."
        Exit Function
    End If
    
    '   Get the Saved Status
    '
    Dim WasSaved As Boolean
    WasSaved = Item.Saved

    '   Check that we can play with it
    '
    If Not FollowUp_Precheck(NoIMAP:=True, Caller:=Caller) Then Exit Function
    
    '   Show the Follow Up Dialog
    '
    '       If Cancel - Done
    '       If Clear - Clear the Item Follow Up
    '       Else - Update the Item Follow Up
    '
    Dim FlagRequest As String
    Dim ReminderTime As String
    If Not FollowUp_Dialog(Item, FlagRequest, ReminderTime) Then
        Exit Function
    End If
    If FlagRequest = "" Then
        If Not FollowUp_Clear() Then Stop: Exit Function
    Else
        If Not FollowUp_SetSave(Item, FlagRequest, ReminderTime) Then Stop: Exit Function
        Categories_AddCat Item, glbCatFollowUp
    End If
    
    '   If was Saved and is not Saved now - Save
    '
    If (WasSaved And Not Item.Saved) Then Item.Save
    
FollowUp_Set = True
End Function

Public Function FollowUp_Clear(Optional ByVal Caller As String = "FollowUp_Clear") As Boolean
Const ThisProc = "FollowUp_Clear"
FollowUp_Clear = False

    '   Get the Active Object
    '
    Dim Item As Object
    If Not Misc_GetActiveItem(Item, Caller) Then Exit Function

    '   Get Saved Status
    '
    Dim WasSaved As Boolean
    WasSaved = Item.Saved

    '   Check that we can play with it
    '
    If Not FollowUp_Precheck(NoIMAP:=True, Caller:=Caller) Then Exit Function
    
    '   Do a ClearTaskFlag
    '
    '       This clears all the Task stuff and the Reminder.
    '
    On Error Resume Next
    
        Item.ClearTaskFlag
        Select Case Err.Number
            Case glbError_None, glbError_PropertyNotFound
            Case Else
                Stop: Exit Function
        End Select
    
    On Error GoTo 0
    
    '   Remove any Follow Up Cat
    '
    Categories_RemoveCat Item, glbCatFollowUp
    
    '   If was Saved and is not Saved now - Save
    '
    If (WasSaved And Not Item.Saved) Then Item.Save
    
FollowUp_Clear = True
End Function

'   PreCheck for a Set/Clear Follow Up with Reminder on the Active Window
'
Public Function FollowUp_Precheck( _
    ByVal NoIMAP As Boolean, _
    ByVal Caller As String) _
    As Boolean
Const ThisProc = "FollowUp_Precheck"
FollowUp_Precheck = False

    '   Active Window must be an Inspector or an Explorer with only one item selected.
    '
    Dim Item As Object
    If Not Misc_GetActiveItem(Item, Caller) Then Exit Function
    
    '   Get the Item Subject if it has one
    '
    Dim Subject As String
    On Error Resume Next
        Subject = Item.Subject
        If Err.Number <> glbError_None Then Subject = ""
    On Error GoTo 0
    
    '   If No IMAP allowed and Active Window is IMAP - Error Exit
    '
    If NoIMAP Then
        If IMAP_ActiveWindowIsIMAP() Then
            Msg_Box _
                Proc:=Caller, _
                Step:="IMAP Check", _
                Subject:=Subject, _
                Text:="Can not Set/Clear a Follow Up with Reminder for an Item in an IMAP Folder."
            Exit Function
        End If
    End If
    
    '   Check for Classes that don't work or will have problems
    '
    If Not FollowUp_CanSet(Item) Then
        Msg_Box _
            Proc:=Caller, _
            Step:="CanSet", _
            Subject:=Subject, _
            Text:="Can not Set/Clear a Follow Up with Reminder on " & _
                   "Item Type '" & Misc_ItemTypeName(Item) & "'. (It don't support it)."
        Exit Function
    End If
    
FollowUp_Precheck = True
End Function

'   Show the Follow Up form and get the values.
'
'       Returns FALSE if Cancel.
'       Returns FlagRequest = "" if Clear.
'
Private Function FollowUp_Dialog( _
    ByVal Item As Object, _
    ByRef FlagRequest As String, _
    ByRef ReminderTime As String _
    ) As Boolean
FollowUp_Dialog = False

    Dim FollowUpForm As FollowUp_Form
    Set FollowUpForm = New FollowUp_Form
    
    '   Populate the Form
    '
    With FollowUpForm
    
        '   Using passed param
        '
        If (FlagRequest <> "") Then
        
            .Title = FlagRequest
            .DateTime = ReminderTime
            
        '   Item already has a FollowUp - Get current values
        '
        ElseIf Item.IsMarkedAsTask Then
        
            '   Why sFlagRequest? See notes on Misc_OLGetProperty.
            '   Why Property instead of Item.FlagRequest? Post doesn't expose.
            '
            ' .Title = Item.FlagRequest
            Dim sFlagRequest As String
            If Not Misc_OLGetProperty(Item, glbPropTag_FlagRequest, sFlagRequest) Then Stop: Exit Function
            .Title = sFlagRequest
            
            .DateTime = Item.ReminderTime
            
        '   New FollowUp - Date = Tomorrow at 5PM
        '
        Else
        
            .Title = ""
            .DateTime = DateAdd("h", 17, DateAdd("d", 1, Date))
        
        End If
        
        '   Show the Form and wait for Hide.
        '
        '       .Show {vbModal, vbModeless} - Won't take named arguments.
        '
        .Show vbModal
        
        '   Process return from the form
        '
        If .Canceled Then
            Exit Function
        ElseIf .Cleared Then
            FlagRequest = ""
            ReminderTime = ""
        Else
            FlagRequest = .TextTitle.value
            ReminderTime = .DateTime
        End If
    
    End With
    
FollowUp_Dialog = True
End Function

'   Set the Item FollowUp values
'
Public Function FollowUp_SetSave( _
    ByVal Item As Object, _
    ByVal FlagRequest As String, _
    ByVal ReminderTime As String _
    ) As Boolean
FollowUp_SetSave = False

    '   Update the Item
    '
    With Item
    
        .MarkAsTask Outlook.olMarkNoDate
        
        ' .IsMarkedAsTask = True                ' Done by MarkAsTask
        ' .FlagIcon = Outlook.olRedFlagIcon     ' Done by MarkAsTask
        ' .FlagRequest = "Follow Up"            ' Done by MarkAsTask
        ' .FlagStatus = Outlook.olFlagMarked    ' Done by MarkAsTask
        
        '   SPOS FlagRequest
        '
        '   Sometimes he shows my text, sometimes "Follow Up" and sometimes "Follow Up flag text
        '   has been hidden". All of which seem to occur at radmon.
        '   See https://www.msoutlook.info/question/254
        '
        '   Everybody hates my Custom Forms.
        '
        '       In this case it's not really Stupid's fault.
        '       They are derived from PostItem which doesn't expose the
        '       FlagRequest property. So we have to deal with it the hard way.
        '
        ' .FlagRequest = FlagRequest
        '
        If Not Misc_OLSetProperty(Item, glbPropTag_FlagRequest, FlagRequest) Then Stop: Exit Function
        
        .ReminderSet = True
        
        '   CDate required
        '
        .ReminderTime = CDate(ReminderTime)
        
    End With
    
FollowUp_SetSave = True
End Function

'   Get  FollowUp FlagRequest
'
Public Function FollowUp_GetFlagRequest( _
    ByVal Item As Object, _
    ByRef FlagRequest As String _
    ) As Boolean
FollowUp_GetFlagRequest = False

    '   Try as a built in Property and then using Property Accessor
    '
    If Not Item.ItemProperties.Item("FlagRequest") Is Nothing Then
        FlagRequest = Item.FlagRequest
    Else
        If Not Misc_OLGetProperty(Item, glbPropTag_FlagRequest, FlagRequest) Then Stop: Exit Function
    End If
    
FollowUp_GetFlagRequest = True
End Function


'   Get FollowUp values and stash them in Billing Info
'
Public Function FollowUp_SetStash(ByVal Item As Object) As Boolean
FollowUp_SetStash = False

    '   Get the Saved Status
    '
    Dim WasSaved As Boolean
    WasSaved = Item.Saved
    
    '   Get any previously Stashed vales
    '
    Dim FlagRequest As String
    Dim ReminderTime As String
    If Not Mail_BIGet(Item, glbBIInx_FlagRequest, FlagRequest) Then Stop: Exit Function
    If Not Mail_BIGet(Item, glbBIInx_ReminderTime, ReminderTime) Then Stop: Exit Function

    '   Show the Dialog, get values, stash them
    '
    If Not FollowUp_Dialog(Item, FlagRequest, ReminderTime) Then Exit Function
    If Not Mail_BISet(Item, glbBIInx_FlagRequest, FlagRequest) Then Stop: Exit Function
    If Not Mail_BISet(Item, glbBIInx_ReminderTime, ReminderTime) Then Stop: Exit Function
    
    '   If was Saved and is not Saved now - Save
    '
    If (WasSaved And Not Item.Saved) Then Item.Save
    
FollowUp_SetStash = True
End Function

'   Does and Object support FollowUp?
'
Private Function FollowUp_CanSet(ByVal Item As Object) As Boolean
FollowUp_CanSet = False

    Dim Dummy As Variant
    On Error Resume Next
    
        '   If it doesn't have ... - No Can Do
        '
        Dummy = Item.IsMarkedAsTask
        If Err.Number <> glbError_None Then Exit Function
        
        Dummy = Item.ReminderSet
        If Err.Number <> glbError_None Then Exit Function
    
    On Error GoTo 0

    '   Why Property instead of Item.FlagRequest? Post doesn't expose.
    '
    ' Dummy = Item.FlagRequest
    ' If Err.Number <> glbError_None Then Exit Function
    '
    If Not Misc_OLGetProperty(Item, glbPropTag_FlagRequest, Dummy) Then Exit Function

FollowUp_CanSet = True
End Function

