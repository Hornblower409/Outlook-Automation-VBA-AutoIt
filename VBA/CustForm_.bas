Attribute VB_Name = "CustForm_"
Option Explicit
Option Private Module

' =====================================================================
'   Custom Forms Table
' =====================================================================

    Private Const CustForm_TypeTable As String = _
        "   " & _
        "   Class                           |   Type                        |    Name               " & vbLf & _
        "   " & _
            glbCustForm_Card & "            |" & gblCustFormType_Card & "   |    Card               " & vbLf & _
            glbCustForm_WipProject & "      |" & gblCustFormType_WIP & "    |    WIP Project        " & vbLf & _
            glbCustForm_WipActivity & "     |" & gblCustFormType_WIP & "    |    WIP Activity       " & vbLf & _
        "'"
        
        Private Const CustForm_TypeTableColClass    As Long = 0
        Private Const CustForm_TypeTableColType     As Long = 1
        Private Const CustForm_TypeTableColName     As Long = 2
        '

' =====================================================================
'   Custom Forms - Table Lookup
' =====================================================================

'   Is a MessageClass a Custom Form?
'
Public Function CustForm_ClassIsCustom(ByVal MessageClass As String) As Boolean
    CustForm_ClassIsCustom = Tbl_TableConstExist(CustForm_TypeTable, MessageClass)
End Function

'   Get a Custom Form Type or "" for a MessageClass
'
Public Function CustForm_ClassType(ByVal MessageClass As String) As String
    Tbl_TableConstFind CustForm_TypeTable, MessageClass, CustForm_TypeTableColType, CustForm_ClassType
End Function

'   Get a Custom Form Name or "" for a MessageClass
'
Public Function CustForm_ClassName(ByVal MessageClass As String) As String
    Tbl_TableConstFind CustForm_TypeTable, MessageClass, CustForm_TypeTableColName, CustForm_ClassName
End Function

' =====================================================================
'   Custom Form - Open
' =====================================================================

Public Function CustForm_OpenEvent(ByVal oForm As Outlook.PostItem) As Boolean
CustForm_OpenEvent = False

    '   Exit Event Scope - Come back at CustForm_Open
    '
    If Not Event_ExitEventScope(oForm, glbExitEventScope_goCustFormOpen) Then Stop: Exit Function
    
CustForm_OpenEvent = True
End Function

Public Function CustForm_Open(ByVal oForm As Outlook.PostItem) As Boolean
CustForm_Open = False

    '   Get the Form Item's InspectorRec
    '
    Dim InspectorRec As Inspector_InspectorRec
    If Not Inspector_RecFromItem(oForm, InspectorRec) Then Stop: Exit Function
    
    With InspectorRec
    
        '   If in DesignMode - Done
        '
        If Ribbon_DesignModeActive(.oInspector) Then CustForm_Open = True: Exit Function
    
        '   Branch based on New or Existing Item
        '
        If .oItem.EntryId = "" Then
            If Not CustForm_OpenNew(InspectorRec) Then Exit Function
        Else
            If Not CustForm_OpenExisting(InspectorRec) Then Exit Function
        End If
    
        '   Branch on MessageClass
        '
        Select Case .oItem.MessageClass
            Case glbCustForm_Card
                If Not CustForm_CardOpen(InspectorRec.oItem) Then Stop: Exit Function
            Case glbCustForm_WipProject
                If Not CustForm_WipProjOpen(InspectorRec.oItem) Then Stop: Exit Function
            Case glbCustForm_WipActivity
                '   Continue
            Case Else
                Stop: Exit Function ' Oops
        End Select

    End With
    
CustForm_Open = True
End Function

Public Function CustForm_OpenNew(ByRef InspectorRec As Inspector_InspectorRec) As Boolean
CustForm_OpenNew = False

    '   Add/Appy Hot Rod Normal Style
    '
    If Not Format_HotRodStyle(InspectorRec.oItem) Then Stop: Exit Function
    
CustForm_OpenNew = True
End Function

Public Function CustForm_OpenExisting(ByRef InspectorRec As Inspector_InspectorRec) As Boolean
Const ThisProc = "CustForm_OpenExisting"
CustForm_OpenExisting = False

    With InspectorRec
    
        '   Wait for the Form to finish Activate
        '
        Dim Index As Long
        For Index = 1 To 30
        
            If Not CustForm_SetFocus(.oInspector) Then Stop: Exit Function
            If Ribbon_EditModeActivate(.oInspector) Then Exit For
            
            Sleep 100
            DoEvents
        
        Next Index
        
        '   If won't Activate Edit Mode - Msg and Done
        '
        If Not Ribbon_EditModeActive(.oInspector) Then
        
            Msg_Box Proc:=ThisProc, Step:="Ribbon_EditModeActivate", Text:= _
                "Inspector: " & .oInspector.Caption & glbBlankLine & _
                "Did not enable Edit Mode."
            Exit Function
            
        End If
    
    End With
    
CustForm_OpenExisting = True
End Function

'   Set the focus for a Custom Form Inspector to the Message Body
'
Public Function CustForm_SetFocus(ByVal oInspector As Outlook.Inspector) As Boolean
CustForm_SetFocus = False

    '   WTF? - Everyone hates Custom Forms.
    '
    '   When opening from either an Outlook Explorer or X1 - If the Message Body is the first field
    '   in the Tab Order, when it switches to Edit Mode the cursor is not visable and it won't accept
    '   input. So I make the Message Body last in the Tab Order and set focus on it when switching to Edit Mode.
    '
    '   If opened from X1 - The Inspector won't report as being in Edit Mode (even though
    '   it really is) until after I've put focus on the Message Body.
    
    '   Set focus on the "Message" control on the first page of a Custom Form
    '
    oInspector.ModifiedFormPages.Item(1).Controls.Item("Message").SetFocus

CustForm_SetFocus = True
End Function

' =====================================================================
'   Custom Form - Response
' =====================================================================

Public Function CustForm_Response(ByVal Original As Object, ByRef Response As Object, ByRef ResponseType As Integer) As Boolean
CustForm_Response = False

    '   If not a Custom Form - done
    '   Get the Form Type
    '
    '       SPOS - Response is not fully formed yet. Have to look at Original for some values (e.g. MessageClass)
    '
    If Not CustForm_ClassIsCustom(Original.MessageClass) Then CustForm_Response = True: Exit Function
    Dim FormType As String: FormType = CustForm_ClassType(Original.MessageClass): If FormType = "" Then Stop: Exit Function
    
    '   Clear the TO field
    '
    '       Recipients is a Collection. Can't use a "Forward For".
    '
    Dim iX As Long
    For iX = Response.Recipients.count To 1 Step -1
        Response.Recipients.Remove (iX)
    Next iX
    
    '   2025-07-12 - Reset ExpiryTime to None
    '   - So any reply from an Outlook client won't echo back my glbDateExpiresFlag
    '
    Response.ExpiryTime = glbDateNone
        
    '   Empty the InternetReplyID so the Response will be treated
    '   as an Original by downstream processes (e.g. Cleanup)
    '
    If Not Misc_OLSetProperty(Response, glbPropTag_InternetReplyID, "") Then Stop: Exit Function

    '   ReplyAll -> Reply
    '
    '       Because Stupid doesn't put all the needed Message Recepient fields on a Post
    '
    If ResponseType = glbResponse_ReplyAll Then ResponseType = glbResponse_Reply
    
    '   If a WIP form remove my WIPID "(XXXNNN)" from the end of the Subject
    '
    If FormType = gblCustFormType_WIP Then Response.Subject = Left(Response.Subject, Len(Response.Subject) - 9)
    
    '   If not a Forward - done
    '
    If Not ResponseType = glbResponse_Forward Then CustForm_Response = True: Exit Function
    
    '   Build the Reply Header Para Text
    '
        Dim HText As String
        
        '   SPOS - Stupid will just ignore any Chr(7) in inserted text. No error, nothing in the docs.
        '   He just doesn't insert the char. (Probably because he needs them as Table Row End Markers).
        '   So don't try using them as your Line Start Anchor.
        '
        Dim LineAnchorChr As String: LineAnchorChr = ChrW(glbUnicode_LineAnchor)
        
        '   Subject: {Subject}
        '
        HText = LineAnchorChr & "Subject: " & Response.Subject & vbVerticalTab
        
        '   If a WIP Custom Form add additional fields
        '
        If FormType = gblCustFormType_WIP Then
        
            '   Description: {Description}
            '
                If UserProp_Get(Original, "WIPDescription") <> "" Then
                    HText = HText & LineAnchorChr & "Description: " & UserProp_Get(Original, "WIPDescription") & vbVerticalTab
                End If
            
            '   Status: {Status} - {StatusUpdate}
            '
            HText = HText & LineAnchorChr & "Status: " & UserProp_Get(Original, "WIPStatus")
            If UserProp_Get(Original, "WIPStatusUpdate") <> "" Then
                HText = HText & " - " & UserProp_Get(Original, "WIPStatusUpdate")
            End If
            HText = HText & vbVerticalTab
            
            '   Priority: {Priority} ({Form Name})
            '
            HText = HText & LineAnchorChr & "Priority: " & UserProp_Get(Original, "WIPPriority")
            HText = HText & " (" & CustForm_ClassName(Original.MessageClass) & ")" & vbVerticalTab
        
        End If
        
        '   Last Updated: {LastModificationTime}
        '
        HText = HText & LineAnchorChr & "Last Updated: " & Format(Original.LastModificationTime, "yyyy-mm-dd")
        
    '   Setup for Word work
    '
    Dim wDoc As Word.Document
    Set wDoc = Response.GetInspector.WordEditor
    
    '   Delete the RepSep para (3) to get rid of the HTML Top Border (and the text)
    '   Put a horizontal line in the empty para (now 3) that was below the RepSep para
    '
    wDoc.Paragraphs.Item(3).Range.Text = ""
    wDoc.Paragraphs.Item(3).Range.InlineShapes.AddHorizontalLineStandard
    
    '   Add a new para (4) below the horizontal line
    '   Fill it with the Header Para Text + an empty para
    '
    wDoc.Paragraphs.Item(3).Range.InsertParagraphAfter
    wDoc.Paragraphs.Item(4).Range.Text = HText & vbCr & vbCr
    Dim HPara As Word.Paragraph: Set HPara = wDoc.Paragraphs.Item(4)
    
    '   Set the Header Para font
    '
    With HPara.Range.Font
    
        .Name = "Tahoma"
        .Size = 10
        
    End With
    
    '   Bold the first "Xxxx:" of each line of the Header Para
    '
    Dim wSearch As Word.Range
    Set wSearch = Word_FindDefault(HPara.Range.Duplicate)
    With wSearch.Find
    
        .MatchWildcards = True
        .Text = LineAnchorChr & "*:"
        .Replacement.Font.Bold = True
        .Execute Replace:=wdReplaceAll
    
    End With
    
    '   Remove the Line Start Anchors
    '
    Set wSearch = Word_FindDefault(HPara.Range.Duplicate)
    With wSearch.Find
    
        .Text = LineAnchorChr
        .Replacement.Text = ""
        .Execute Replace:=wdReplaceAll
    
    End With

CustForm_Response = True
End Function

' =====================================================================
'   Custom Form - Controls
' =====================================================================

'   Search any "ModifiedFormPages" (Custom Form Pages) for a Control by Name
'
Public Function CustForm_ControlByName(ByVal oForm As Object, ByVal ControlName As String, ByRef oControl As MSForms.Control) As Boolean
CustForm_ControlByName = False

    '   Get the Form's Inspector
    '
    If Not Inspector_ItemInspectorExist(oForm) Then Stop: Exit Function
    Dim oInspector As Outlook.Inspector
    Set oInspector = oForm.GetInspector
    
    '   Get the Inspector's ModifiedFormPages
    '
    Dim oPages As Outlook.Pages
    Set oPages = oInspector.ModifiedFormPages
    If oPages Is Nothing Then Stop: Exit Function
    
    '   Search the ModifiedFormPages for the Control
    '
    Dim oPage As MSForms.UserForm
    Dim PageInx As Long
    For PageInx = 1 To oPages.count
        
        Set oPage = oPages(PageInx)
        Dim oControls As MSForms.Controls
        Set oControls = oPage.Controls
        
        On Error Resume Next
            Set oControl = oControls.Item(ControlName)
            Select Case Err.Number
                Case glbError_None
                    CustForm_ControlByName = True
                    Exit Function
                Case glbError_ItemNotInCollection
                    '   Continue
                Case Else
                    Stop: Exit Function
            End Select
        On Error GoTo 0

    Next PageInx

End Function

' =====================================================================
'   Custom Form - Prop Change
' =====================================================================

Public Function CustForm_PropChange(ByVal oForm As Object, ByVal IsStandard As Boolean, ByVal PropName As String) As Boolean
CustForm_PropChange = False

    '   Branch on MessageClass
    '
    Select Case oForm.MessageClass
        Case glbCustForm_Card
            If Not CustForm_CardPropChange(oForm, IsStandard, PropName) Then Stop: Exit Function
        Case glbCustForm_WipProject
            If Not CustForm_WipProjPropChange(oForm, IsStandard, PropName) Then Stop: Exit Function
        Case glbCustForm_WipActivity
            '   Continue
        Case Else
            Stop: Exit Function ' Oops
    End Select

CustForm_PropChange = True
End Function

' =====================================================================
'   Custom Form - Form Specific Custom Actions
' =====================================================================

Public Function CustForm_CustAction(ByVal oForm As Outlook.PostItem, ByVal Action As String) As Boolean
CustForm_CustAction = False

    '   Branch on MessageClass
    '
    Select Case oForm.MessageClass
        Case glbCustForm_Card
            If Not CustForm_CardCustAction(oForm, Action) Then Stop: Exit Function
        Case glbCustForm_WipProject
            If Not CustForm_WipProjCustAction(oForm, Action) Then Stop: Exit Function
        Case glbCustForm_WipActivity
            '   Continue
        Case Else
            Stop: Exit Function ' Oops
    End Select

CustForm_CustAction = True
End Function

Public Function CustForm_CustActionEnabled(ByVal oPost As Outlook.PostItem, ByVal Action As String) As Boolean
CustForm_CustActionEnabled = False

    Dim oActions As Outlook.Actions
    Set oActions = oPost.Actions
    
    Dim oAction As Outlook.Action
    Set oAction = oActions.Item(Action)
    
    If oAction Is Nothing Then Exit Function
    If Not oAction.Enabled Then Exit Function

CustForm_CustActionEnabled = True
End Function

' =====================================================================
'   Custom Form - Write
' =====================================================================

Public Function CustForm_Write(ByVal oForm As Outlook.PostItem) As Boolean
CustForm_Write = False

    '   If oForm not Open in an Inspector - Bail
    '
    If Not Inspector_ItemInspectorExist(oForm) Then
        CustForm_Write = True
        Exit Function
    End If
    
    '   Get the oForm InspectorRec
    '
    Dim InspectorRec As Inspector_InspectorRec
    If Not Inspector_RecFromItem(oForm, InspectorRec) Then Stop: Exit Function
    
    With InspectorRec
    
        '   If the Inspector is in DesignMode - Bail
        '
        If Ribbon_DesignModeActive(.oInspector) Then
            CustForm_Write = True
            Exit Function
        End If
        
        '   Branch on MessageClass
        '
        Select Case oForm.MessageClass
            Case glbCustForm_WipProject
                If Not CustForm_WipProjWrite(oForm) Then Exit Function
            Case Else
                '   Continue
        End Select

    End With
    
CustForm_Write = True
End Function

' =====================================================================
'   Custom Form - Command Button Click
' =====================================================================

Public Function CustForm_ClickCmdButton(ByVal oForm As Outlook.PostItem, ByVal oCmdButton As MSForms.CommandButton) As Boolean
CustForm_ClickCmdButton = False

    '   Branch on MessageClass
    '
    Select Case oForm.MessageClass
        Case glbCustForm_Card
            If Not CustForm_CardClickCmdButton(oForm, oCmdButton) Then Stop: Exit Function
        Case glbCustForm_WipProject
            If Not CustForm_WipProjClickCmdButton(oForm, oCmdButton) Then Stop: Exit Function
        Case Else
            '   Continue
    End Select

CustForm_ClickCmdButton = True
End Function

' =====================================================================
'   Custom Form - Manuals
' =====================================================================

Public Sub CustForm_MessageClassChange()

    ' ************************************************************
    Dim sFolderPath As String: sFolderPath = "\\Cards\Cards"
    Dim sOldMsgClass As String: sOldMsgClass = "IPM.Post"
    Dim sNewMsgClass As String: sNewMsgClass = "IPM.Post.CardV3"
    ' ************************************************************

    Dim oFolder As Outlook.Folder
    Set oFolder = Folders_Path(sFolderPath)
    If oFolder Is Nothing Then Stop: Exit Sub

    '   Loop thru all Items
    '
    Dim LoopCounter As Long
    Dim ChangedCounter As Long
    Dim oItem As Object
    For Each oItem In oFolder.Items

        If oItem.MessageClass = sOldMsgClass Then
            oItem.MessageClass = sNewMsgClass
            oItem.Save
            ChangedCounter = ChangedCounter + 1
        End If

        LoopCounter = LoopCounter + 1
        If (LoopCounter Mod 100) = 0 Then Debug.Print LoopCounter

    Next oItem

    Debug.Print "Changed: " & ChangedCounter

End Sub

