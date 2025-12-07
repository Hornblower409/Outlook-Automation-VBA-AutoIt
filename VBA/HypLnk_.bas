Attribute VB_Name = "HypLnk_"
Option Explicit
Option Private Module

' =====================================================================
'   HyperLink
' =====================================================================

'   Put a Hyperlink to the Active Item on the Clipboard
'
Public Sub HypLnk_HyperlinkGet()
Const ThisProc = "HypLnk_HyperlinkGet"

    Dim oTarget As Object
    If Not Misc_GetActiveItem(oTarget, ThisProc, Saved:=True) Then Exit Sub
    If Not HypLnk_HyperlinkClip(oTarget, Caller:=ThisProc) Then Exit Sub
    
End Sub

'   Put a formatted Hyperlink to oTarget on the Clipboard
'
Public Function HypLnk_HyperlinkClip(ByVal oTarget As Object, ByVal Caller As String) As Boolean
HypLnk_HyperlinkClip = False

    '   Can not operate on an unsaved Item
    '
    If Not oTarget.Saved Then Stop: GoTo Exit_Function
    
    ' ---------------------------------------------------------------------
    '   Get Default Link data
    ' ---------------------------------------------------------------------
    
    Dim sSubject As String
    If Not oTarget.ItemProperties.Item("Subject") Is Nothing Then sSubject = oTarget.Subject
    
    Dim sLinkID As String
    sLinkID = oTarget.EntryId
    
    Dim sLinkName As String
    sLinkName = sSubject
    
    Dim sTypeName As String
    sTypeName = TypeName(oTarget)
    
    '   sLinkType is the TypeName without the "Item" suffix
    '
    Dim sLinkType As String
    sLinkType = Left(sTypeName, Len(sTypeName) - Len("Item"))
        
    ' ---------------------------------------------------------------------
    '   Update the Default Link data based on TypeName(oTarget)
    '   And if it is one of my Custom Forms.
    '---------------------------------------------------------------------
    
    Select Case sTypeName
    
        Case "MailItem", "TaskItem", "NoteItem", "AppointmentItem", "PostItem"
        
            '   All Default
            
        Case "ContactItem"
        
            '   SPOS - ContactItem.CompanyAndFullName is NOT what it claims to be.
            '   Have to do the combination myself.
            '
            Dim Delimeter As String
            If oTarget.FullName <> "" And oTarget.CompanyName <> "" Then Delimeter = " - "
            sLinkName = Join(Array(oTarget.FullName, oTarget.CompanyName), Delimeter)
            
        Case "MeetingItem"
        
            '   SPOS - It's not a "Meeting" stupid, it's an Invite
            '
            sLinkType = "Invite"
    
        Case Else
        
            Msg_Box Proc:=Caller, Step:="Select Case sTypeName", _
                Subject:=sSubject, _
                Text:="Do not know how to create a Hyperlink to this type of Item." & glbBlankLine & _
                      "TypeName(oTarget): " & sTypeName
            GoTo Exit_Function
        
    End Select
    
    '   If one of my Custom Forms - LinkType = My Custom Form Name
    '
    If CustForm_ClassIsCustom(oTarget.MessageClass) Then sLinkType = CustForm_ClassName(oTarget.MessageClass)
    
    ' ---------------------------------------------------------------------
    '   HotRodGUID Group UserProps
    ' ---------------------------------------------------------------------
    
    '   If this Item already has HotRodGUID User Props - Check them.
    '   Else - Gen New and Add Them
    '
    Dim HotRodGUIDRec As HotRodGUID_HotRodGUIDRec
    If HotRodGUID_Get(oTarget, HotRodGUIDRec) Then
    
        With HotRodGUIDRec
        
            If Not Session.CompareEntryIDs(.EntryId, oTarget.EntryId) Then
                Msg_Box Proc:=Caller, Step:="Cross Check HotRodGUID EntryId", _
                    Subject:=sSubject, _
                    Text:="HotRodGUID EntryId does not match this Item's EntryId. (Item has been moved or copied)" & glbBlankLine & _
                          "Item.EntryId: " & oTarget.EntryId & vbNewLine & _
                          "HotRod GUID: " & .GUID & vbNewLine & _
                          "HotRod EntryId: " & .EntryId & vbNewLine & _
                          "HotRod EntryIdMod: " & .EntryIdMod
                GoTo Exit_Function
            End If
            
        End With
        
    Else
    
        If Not HotRodGUID_Add(HotRodGUIDRec) Then Stop: GoTo Exit_Function
        
    End If
    
    ' ---------------------------------------------------------------------
    '   Build the Hyperlink
    ' ---------------------------------------------------------------------
    
    Dim LinkAddress As String
    LinkAddress = "outlook:" & sLinkID
    
    Dim LinkScreenTip As String
    LinkScreenTip = glbHotRodGUIDLabel & HotRodGUIDRec.GUID
    
    Dim LinkTextToDisplay As String
    LinkTextToDisplay = sLinkType & ": " & sLinkName
    
    '   Create a new Post
    '
    Dim oPost As Outlook.PostItem
    Set oPost = Application.CreateItem(Outlook.olPostItem)
    
    '   Get the new Post InspectorRec and wDoc
    '
    Dim InspectorRec As Inspector_InspectorRec
    If Not Inspector_RecFromItem(oPost, InspectorRec) Then Stop: GoTo Exit_Function
    
    Dim wDoc As Word.Document
    Set wDoc = InspectorRec.oInspector.WordEditor
    If wDoc Is Nothing Then Stop: GoTo Exit_Function
    
    '   Apply HotRod Normal Font and Tabs
    '
    If Not Format_SetHotRodNormalFont(wDoc.Range.Font) Then Stop: GoTo Exit_Function
    If Not Format_SetTabs(wDoc.Range.ParagraphFormat) Then Stop: GoTo Exit_Function
    
    '   Build the Hyperlink in the wDoc
    '
    wDoc.Hyperlinks.Add _
        Anchor:=wDoc.Range(0, 0), _
        Address:=LinkAddress, _
        SubAddress:="", _
        ScreenTip:=LinkScreenTip, _
        TextToDisplay:=LinkTextToDisplay
    
    '   Cleanup the Hyperlink formatting
    '
    With wDoc.Range.Font
        .Name = "Courier New"
        .Size = 10
    End With
    
    '   Put the Plain and Fancy links on the Clipboard
    '
    Misc_ClipSet LinkAddress
    wDoc.Range.Copy
    Misc_ClipSetWait
    
HypLnk_HyperlinkClip = True
Exit_Function:

    If Not Inspector_RecInspectorCloseIfNew(InspectorRec) Then Stop: Exit Function
    
End Function

Public Sub HypLnk_HyperlinkCheckFolder_Manual()
Const ThisProc = "HypLnk_HyperlinkCheckFolder_Manual"

    ' ************************************************************
    Dim sFolderPath As String: sFolderPath = "\\Cards\Cards"
    ' ************************************************************

    Debug_HotRodLog Proc:=ThisProc, Step:="Entry", Text:=" -----  Log Start  ----- "

    Dim oFolder As Outlook.Folder
    Set oFolder = Folders_Path(sFolderPath)
    If oFolder Is Nothing Then Stop: Exit Sub

    '   Loop thru all Items
    '
    Dim ErrorCounter As Long
    Dim LoopCounter As Long
    Dim oItem As Object
    For Each oItem In oFolder.Items

        If Not HypLnk_HyperlinkCheckItem(oItem, ErrorCounter) Then Stop: Exit Sub
        
        LoopCounter = LoopCounter + 1
        If (LoopCounter Mod 100) = 0 Then Debug.Print LoopCounter

    Next oItem

    Debug_HotRodLog Proc:=ThisProc, Step:="Exit", Text:="Errors", P1Name:="ErrorCounter", P1Value:=ErrorCounter
    Debug_HotRodLog Proc:=ThisProc, Step:="Exit", Text:=" -----  Log End  ----- "
    
    If ErrorCounter > 0 Then
        Msg_Box Proc:=ThisProc, Step:="Exit", _
        Text:="ErrorCounter: " & ErrorCounter & glbBlankLine & _
              "See: " & glbHotRodLogFile
    End If
    
End Sub

Public Function HypLnk_HyperlinkCheckItem(ByVal oSource As Object, ByRef ErrorCounter As Long) As Boolean
Const ThisProc = "HypLnk_HyperlinkCheckItem"
HypLnk_HyperlinkCheckItem = False

    ' =====================================================================
    '   Source
    ' =====================================================================
    
    '   Source must be Saved
    '
    If Not oSource.Saved Then Stop: GoTo Exit_Function
    
    '   Get the Source Subject
    '
    Dim SourceSubject As String
    If Not (oSource.ItemProperties.Item("Subject") Is Nothing) Then SourceSubject = oSource.Subject
    
    '   Get the Source HotRodGUID User Props (if any).
    '   If it has them - Verify.
    '
    Dim SourceHotRodGUIDRec As HotRodGUID_HotRodGUIDRec
    If HotRodGUID_Get(oSource, SourceHotRodGUIDRec) Then
    
        With SourceHotRodGUIDRec
        
            If Not Session.CompareEntryIDs(.EntryId, oSource.EntryId) Then
            
                Debug_HotRodLog Proc:=ThisProc, Step:="Source HotRodGUID Check", Subject:=SourceSubject, _
                                Text:=" Source EntryId does not match the HotRodGUID EntryId (Item has been moved or copied)", _
                                P1Name:="Source EntryId", P1Value:=oSource.EntryId, _
                                P2Name:="Source HotRodGUID", P2Value:=.GUID, _
                                P3Name:="Source HotRodGUID EntryId", P3Value:=.EntryId, _
                                P4Name:="Source HotRodGUID EntryIdMod", P4Value:=.EntryIdMod
                                
                ErrorCounter = ErrorCounter + 1
                
            End If
            
        End With
        
    End If
    
    '   Get the Source InspectorRec and wDoc
    '
    Dim InspectorRec As Inspector_InspectorRec
    If Not Inspector_RecFromItem(oSource, InspectorRec) Then Stop: GoTo Exit_Function
    Dim wDoc As Word.Document
    Set wDoc = InspectorRec.oInspector.WordEditor
    
    '   Get the Hyperlinks in the Item
    '
    Dim oHyperlinks As Word.Hyperlinks
    Set oHyperlinks = wDoc.Hyperlinks
    
    '   Loop thru all the Hyperlinks
    '
    Dim oHyperlink As Word.Hyperlink
    For Each oHyperlink In oHyperlinks: Do
    
        DoEvents
        Sleep 25
        
        '   Extract the Hyperlink fields
        '
        Dim sAddress As String
        Dim sScreenTip As String
        Dim sTextToDisplay As String
        Dim sText As String
        
        On Error Resume Next
        
            sAddress = oHyperlink.Address
            sScreenTip = oHyperlink.ScreenTip
            sTextToDisplay = oHyperlink.TextToDisplay
            sText = oHyperlink.Range.Text
            
            Select Case Err.Number
            
                Case glbError_None
                    '   Continue
                    
                Case glbError_CommandFailed
                
                    Debug_HotRodLog Proc:=ThisProc, Step:="Access Hyperlink", Subject:=SourceSubject, _
                                    oErr:=Err, _
                                    Text:="glbError_CommandFailed on oHyperlink access."
                                    
                    ErrorCounter = ErrorCounter + 1
                    Exit Do ' Next oHyperlink
                
                Case Else
                    Stop: GoTo Exit_Function
                    
            End Select
            
        On Error GoTo 0
        
        '   If not an Outlook link - Skip it.
        '
        If Not (InStr(1, sAddress, "outlook:", vbTextCompare) = 1) Then Exit Do  ' Next oHyperlink
        
        '   Get the Target EntryId
        '
        Dim sEntryId As String
        sEntryId = Mid(sAddress, Len("outlook:") + 1)
        
        '   Get the Target Item - Error if broken link
        '
        Dim oTarget As Object
        Set oTarget = Misc_GetItemFromID(sEntryId)
        
        If (oTarget Is Nothing) Then
        
            Debug_HotRodLog Proc:=ThisProc, Step:="Link Check", Subject:=SourceSubject, _
                            Text:="Broken Link", _
                            P1Name:="Target Address", P1Value:=sAddress, _
                            P2Name:="Target Text", P2Value:=sText
            ErrorCounter = ErrorCounter + 1
            Exit Do ' Next oHyperlink
            
        End If
        
        ' =====================================================================
        '   Hyperlink Target
        ' =====================================================================
        
        '   Target must be Saved.
        '
        If Not oTarget.Saved Then Stop: GoTo Exit_Function
        
        '   Get the Target Subject
        '
        Dim TargetSubject As String
        If Not (oTarget.ItemProperties.Item("Subject") Is Nothing) Then TargetSubject = oTarget.Subject
    
        '   Get the Target HotRodGUID Props (if any)
        '
        Dim TargetHotRodGUIDRec As HotRodGUID_HotRodGUIDRec
        Dim TargetHasProps As Boolean
        TargetHasProps = HotRodGUID_Get(oTarget, TargetHotRodGUIDRec)
        
        '   Branch on the ScreenTip value in the Source Hyperlink
        '
        Select Case sScreenTip
        
            '   ScreenTip is blank.
            '
            Case ""
            
                '   If Target does not have Props - Add and Notify
                '
                If Not TargetHasProps Then
                
                    If Not HotRodGUID_Add(TargetHotRodGUIDRec) Then Stop: GoTo Exit_Function
                    Debug_HotRodLog Proc:=ThisProc, Step:="Add HotRodGUID Props to Target", Subject:=SourceSubject, _
                                    Text:="HotRodGUID User Props added to Hyperlink Target.", _
                                    P1Name:="Target Address", P1Value:=sAddress, _
                                    P2Name:="Target Text", P2Value:=sText
                    
                End If
                
                '
                '   Populate the Source ScreenTip
                '
                If wDoc.ProtectionType <> wdNoProtection Then wDoc.UnProtect
                oHyperlink.ScreenTip = glbHotRodGUIDLabel & TargetHotRodGUIDRec.GUID
                oSource.Save
                
                Debug_HotRodLog Proc:=ThisProc, Step:="Add ScreenTip", Subject:=SourceSubject, _
                                Text:="Added Target HotRodGUID to Source Hyperlink.", _
                                P1Name:="Target Address", P1Value:=sAddress, _
                                P2Name:="Target Text", P2Value:=sText
            
            '   ScreenTip is not blank.
            '
            Case Else
            
                '   If non-blank ScreenTip doesn't start with HotRodGUIDLabel - Error
                '
                If Not InStr(1, sScreenTip, glbHotRodGUIDLabel) = 1 Then
                
                    Debug_HotRodLog Proc:=ThisProc, Step:="Check ScreenTip has HotRodGUID prefix", Subject:=SourceSubject, _
                                    Text:="Source Hyperlink has a non-blank ScreenTip that does not have a HotRodGUID prefix.", _
                                    P1Name:="Target Address", P1Value:=sAddress, _
                                    P2Name:="Target Text", P2Value:=sText, _
                                    P3Name:="ScreenTip", P3Value:=sScreenTip
                    ErrorCounter = ErrorCounter + 1
                    Exit Do ' Next oHyperlink
                
                End If
                
                '   Get the Target GUID from the ScreenTip
                '
                Dim sScreenTipGUID As String
                sScreenTipGUID = Mid(sScreenTip, Len(glbHotRodGUIDLabel) + 1)
            
                '   If the Target does not have HotRodGUID Props - Error
                '
                If Not TargetHasProps Then
                
                    Debug_HotRodLog Proc:=ThisProc, Step:="Check Target HotRodGUID Props Exist", Subject:=SourceSubject, _
                                    Text:="Source Hyperlink has a HotRodGUID in the ScreenTip but the Target does not have all the HotRodGUID User Props.", _
                                    P1Name:="Target Address", P1Value:=sAddress, _
                                    P2Name:="Target Text", P2Value:=sText, _
                                    P3Name:="ScreenTip", P3Value:=sScreenTip
                    ErrorCounter = ErrorCounter + 1
                    Exit Do ' Next oHyperlink
                
                End If
                
                '   If the ScreenTip GUID doesn't match the Target GUID - Error
                '
                If sScreenTipGUID <> TargetHotRodGUIDRec.GUID Then
                
                    Debug_HotRodLog Proc:=ThisProc, Step:="Check Target HotRodGUID match", Subject:=SourceSubject, _
                                    Text:="Source Hyperlink has a HotRodGUID in a ScreenTip that does not match the Target HotRodGUID.", _
                                    P1Name:="Target Address", P1Value:=sAddress, _
                                    P2Name:="Target Text", P2Value:=sText, _
                                    P3Name:="ScreenTip HotRodGUID", P3Value:=sScreenTipGUID, _
                                    P4Name:="Target HotRodGUID", P4Value:=TargetHotRodGUIDRec.GUID
                    ErrorCounter = ErrorCounter + 1
                    Exit Do ' Next oHyperlink
                
                End If
            
        End Select

    Loop While False: Next oHyperlink
    
HypLnk_HyperlinkCheckItem = True
Exit_Function:

    If Not Inspector_RecInspectorCloseIfNew(InspectorRec) Then Stop: Exit Function
    
End Function

