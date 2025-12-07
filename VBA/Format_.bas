Attribute VB_Name = "Format_"
Option Explicit
Option Private Module

'   2024-01-22 - Set the Font for the current selection
'
'   Called by ProcXeq
'
'       CmdLine(1) = Font Name      Null = No Change. "Consolas"
'       CmdLine(2) = Size           Null = No Change. "12"
'       CmdLine(3) = Bold           Null = No Change. "True", "False"
'       CmdLine(4) = Italic         Null = No Change. "True", "False"
'
Public Function Format_FontSet(ByRef CmdLine() As String) As Boolean
Const ThisProc = "Format_FontSet"
Format_FontSet = False

    '   Arg count check
    '
    If UBound(CmdLine) <> 4 Then
        Msg_Box Proc:=ThisProc, Step:="Arg count check", Text:="ProcXeq command must have exactly four args."
        Exit Function
    End If

    '   Get the current Word Selection
    '
    Dim wrdSelection As Word.Selection
    If Not Format_GetWordSelection(wrdSelection) Then Exit Function

    '   Font Name
    '
    If CmdLine(1) <> "" Then
        wrdSelection.Font.Name = CmdLine(1)
    End If
    
    '   Font Size
    '
    If CmdLine(2) <> "" Then
        wrdSelection.Font.Size = CmdLine(2)
    End If
    
    '   Bold
    '
    If CmdLine(3) <> "" Then
        wrdSelection.Font.Bold = IIf(UCase(CmdLine(3)) = "TRUE", True, False)
    End If
    
    '   Italic
    '
    If CmdLine(4) <> "" Then
        wrdSelection.Font.Italic = IIf(UCase(CmdLine(4)) = "TRUE", True, False)
    End If
    
Format_FontSet = True
End Function

Public Function Format_Tabs() As Boolean
Format_Tabs = False

    Dim wrdSelection As Word.Selection
    If Not Format_GetWordSelection(wrdSelection) Then Exit Function
    If Not Format_SetTabs(wrdSelection.Paragraphs) Then Stop: Exit Function

Format_Tabs = True
End Function

Public Function Format_Normal() As Boolean
Format_Normal = False

    Dim wrdSelection As Word.Selection
    If Not Format_GetWordSelection(wrdSelection) Then Exit Function
    
    '   Force HotRod style and Refresh. Seems to be the only thing that works.
    '
    If Not Format_HotRodNormal(wrdSelection.Font, wrdSelection.ParagraphFormat) Then Stop: Exit Function
    DoEvents
    wrdSelection.Parent.Application.ScreenRefresh
    DoEvents
    
Format_Normal = True
End Function

Public Function Format_Courier10() As Boolean
Format_Courier10 = False

    Dim wrdSelection As Word.Selection
    If Not Format_GetWordSelection(wrdSelection) Then Exit Function
    
    With wrdSelection.Font
        .Name = "Courier New"
        .Size = 10
    End With
    
Format_Courier10 = True
End Function

Public Function Format_Red() As Boolean
Format_Red = False

    Dim wrdSelection As Word.Selection
    If Not Format_GetWordSelection(wrdSelection) Then Exit Function
    
    wrdSelection.Font.Color = wdColorRed
    
Format_Red = True
End Function

Public Function Format_Gray() As Boolean
Format_Gray = False

    Dim wrdSelection As Word.Selection
    If Not Format_GetWordSelection(wrdSelection) Then Exit Function
    
    wrdSelection.Font.Color = wdColorGray40
    wrdSelection.Font.Italic = True

Format_Gray = True
End Function

Public Function Format_BlockQuote() As Boolean
Format_BlockQuote = False

    Dim wrdSelection As Word.Selection
    If Not Format_GetWordSelection(wrdSelection) Then Exit Function
    
    Dim wdParagraphFormat As Word.ParagraphFormat
    Set wdParagraphFormat = wrdSelection.ParagraphFormat
    
    If Not Format_SetBordersNone(wdParagraphFormat) Then Stop: Exit Function
    
    With wdParagraphFormat.Borders(wdBorderLeft)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth300pt
        .Color = wdColorGray15
    End With

    With wdParagraphFormat.Borders
        .DistanceFromLeft = 4
    End With

Format_BlockQuote = True
End Function

'   Bullet Toggle Selection
'
Public Function Format_BulletToggle() As Boolean
Format_BulletToggle = False

    Dim wrdSelection As Word.Selection
    If Not Format_GetWordSelection(wrdSelection) Then Exit Function

    Ribbon_ExecuteMSO ActiveInspector, glbidMSO_BulletsGalleryWord

Format_BulletToggle = True
End Function

'   Show the Font Dialog
'
Public Function Format_FontDialog() As Boolean
Format_FontDialog = False

    Dim wrdSelection As Word.Selection
    If Not Format_GetWordSelection(wrdSelection) Then Exit Function

    Ribbon_ExecuteMSO ActiveInspector, glbidMSO_FontDialog

Format_FontDialog = True
End Function

'   Get the Word Selection Object of the Active Inspector
'
'       Returns False if the Inspector Item is "Locked For editing" (Error 4605)
'       or there is not an Active Inspector.
'
'                                       Must be ByRef
'                                       \/
Public Function Format_GetWordSelection(ByRef wrdSelection As Word.Selection) As Boolean
Const ThisProc = "Format_GetWordSelection"

    '   Default to False (Failed)
    '
    Format_GetWordSelection = False

    '   Make sure we have an Active Inspector with an item loaded
    '
    Dim DummyItem As Object
    If Not Misc_GetActiveItem(DummyItem, InspectorOnly:=True) Then Exit Function
    
    '   Get the Active Inspector Word Editor
    '
    Dim wDoc As Word.Document: Set wDoc = ActiveInspector.WordEditor
    If wDoc Is Nothing Then Exit Function
            
    '   Check the Word Document ProtectionType (Who knew?)
    '
    If wDoc.ProtectionType <> wdNoProtection Then
        Msg_Box Proc:=ThisProc, Text:="The Active Inspector is Locked For Editing (Read Only)."
        Exit Function
    End If
    
    '   Set the Selection and Return True
    '
    Set wrdSelection = wDoc.Windows.Item(1).Selection
    Format_GetWordSelection = True

End Function

'   Format Any New Items to MY Standard
'
'       Called on Open of a New Item
'
'       Add my Hot Rod Normal Style to the Item.
'       Set the first lines of the Item to Hot Rod Normal Style
'
'       Means that any new item will open as unsaved. I don't give a shit.
'       IT SOLVES THE TAB STOP PROBLEM!!
'
Public Function Format_HotRodStyle(ByVal Item As Object) As Boolean
Format_HotRodStyle = True

    '   Get the Item's InspectorRec
    '
    If Not Inspector_ItemInspectorExist(Item) Then Stop: Exit Function
    Dim InspectorRec As Inspector_InspectorRec
    If Not Inspector_RecFromItem(Item, InspectorRec) Then Stop: Exit Function
    
    '   Pick out just the Inspector
    '
    Dim Inspector As Outlook.Inspector
    Set Inspector = InspectorRec.oInspector
    
    ' 2023-01-17 - To handle opening MSG files/attachments. Which have no EntryId but Sent = TRUE.
    '
    '   If it has a Sent property and it is TRUE - done
    '
    If Mail_IsSent(Item) Then Exit Function
    
    '   If it doesn't have a Fancy Editor - done
    '
    If Inspector.WordEditor Is Nothing Then Exit Function
    
    '   If it's Plain Text - done
    '
    If Not (Item.ItemProperties.Item("BodyFormat") Is Nothing) Then
        If Item.BodyFormat = Outlook.olFormatPlain Then Exit Function
    End If
    
    '   Add my Hot Rod Normal Style to the Word Doc
    '
    If Not Format_SetHotRodNormal(Inspector) Then Stop: Exit Function
    
    '   Set the first two paragraph's Style to Hot Rod
    '
    With Inspector.WordEditor.Paragraphs
    
        If .count > 0 Then .Item(1).style = glbHotRodStyle
        If .count > 1 Then .Item(2).style = glbHotRodStyle
        
    End With

End Function

'   Add the Hot Rod Normal Style to an Inspector Doc
'
Public Function Format_SetHotRodNormal(ByVal Inspector As Outlook.Inspector) As Boolean

    Format_SetHotRodNormal = True

    '   Get the Inspector's Word Editor
    '
    Dim wDoc As Word.Document: Set wDoc = Inspector.WordEditor
    If wDoc Is Nothing Then Exit Function
            
    '   If the Word Document is Locked For Editing - we're done
    '
    If wDoc.ProtectionType <> wdNoProtection Then Exit Function

    '   If the doc doesn't have my Hot Rod Style - Add it
    '
    Dim hrnStyle As Word.style
    If Not Word_StyleExists(wDoc.Content, glbHotRodStyle) Then wDoc.Styles.Add glbHotRodStyle, wdStyleTypeParagraph
    
    '   Init or Refresh the HotRod Normal Style for this Doc
    '
    Set hrnStyle = wDoc.Styles.Item(glbHotRodStyle)
    With hrnStyle
    
        .AutomaticallyUpdate = False
        .QuickStyle = True
        
        '   SPOS - This one should win some kind of prize.
        '
        '       Sometime when you set these two properties it thows a Locked For Editing error
        '       even though we checked wDoc.ProtectionType before we got here.
        '
        '       And when called from an Item Open it may cause an unrelated item (e.g. a Post
        '       open in another Inspector) to go Saved = False.
        '
        '       https://stackoverflow.com/questions/18547789/cannot-set-word-styles-base-style-from-c-sharp
        '       "cannot change the base style if [the doc] is invisible" [i.e. before Activate]
        '
        '       Turns out I don't really need them because I have HotRod defined in my NormalEmail.dotm.
        '       So the defaults when we get here are:
        '
        '           BaseStyle is "Normal"
        '           NextParagraphStyle is already HotRod
        '
        '   But I'm leaving this here in case I ever forget and turn them back on.
        '
        ' On Error Resume Next
        '     .BaseStyle = ""
        '     .NextParagraphStyle = hrnStyle
        ' On Error GoTo 0
        
        '   WTF BUG
        '
        '   See https://bettersolutions.com/word/styles/vba-bug-visibility-property.htm
        '   This property actually does the opposite to what it says and if you touch it
        '   all the heading styles in your document will have a list style associated with them.
        '
        '   Just DO NOT TOUCH
        '
        ' .Visibility = True
        
        If Not Format_HotRodNormal(.Font, .ParagraphFormat) Then Stop: Exit Function
        
    End With
    
End Function

Public Function Format_HotRodNormal(ByVal Font As Word.Font, ByVal ParagraphFormat As Word.ParagraphFormat) As Boolean
Format_HotRodNormal = False

        If Not Format_SetHotRodNormalFont(Font) Then Stop: Exit Function
        If Not Format_SetHotRodNormalParagraphFormat(ParagraphFormat) Then Stop: Exit Function
        If Not Format_SetTabs(ParagraphFormat) Then Stop: Exit Function
        If Not Format_SetBordersNone(ParagraphFormat) Then Stop: Exit Function

Format_HotRodNormal = True
End Function

Public Function Format_SetBordersNone(ByVal ParagraphFormat As Word.ParagraphFormat) As Boolean
Format_SetBordersNone = False

    Dim wdBorder As Word.Border
    For Each wdBorder In ParagraphFormat.Borders
        wdBorder.LineStyle = wdLineStyleNone
    Next wdBorder

Format_SetBordersNone = True
End Function

Public Function Format_SetHotRodNormalFont(ByVal Font As Word.Font) As Boolean
Format_SetHotRodNormalFont = False

    With Font
    
        .AllCaps = False
        .Bold = False
        .ColorIndex = wdAuto
        .DoubleStrikeThrough = False
        .Emboss = False
        .Engrave = False
        .Hidden = False
        .Italic = False
        .Name = "Courier New"
        .Outline = False
        .Shadow = False
        .Size = 10
        .SmallCaps = False
        .Strikethrough = False
        .Subscript = False
        .Superscript = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
            
    End With

Format_SetHotRodNormalFont = True
End Function

Public Function Format_SetHotRodNormalParagraphFormat(ByVal ParagraphFormat As Word.ParagraphFormat) As Boolean
Format_SetHotRodNormalParagraphFormat = False

    With ParagraphFormat
    
        .Alignment = wdAlignParagraphLeft
        .CharacterUnitFirstLineIndent = 0
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .FirstLineIndent = 0
        .Hyphenation = False
        .KeepTogether = False
        .KeepWithNext = False
        .LeftIndent = 0
        .LineSpacingRule = wdLineSpaceSingle
        .MirrorIndents = False
        .OutlineLevel = wdOutlineLevelBodyText
        .PageBreakBefore = False
        .RightIndent = 0
        
        '   2025-10-10 - Added Auto Reset
        '
        .SpaceAfterAuto = False
        .SpaceBeforeAuto = False
        
        .SpaceAfter = 0
        .SpaceBefore = 0
        .TabHangingIndent (0)
            
    End With

Format_SetHotRodNormalParagraphFormat = True
End Function

'   Set my standard Tab Stops for an Object (Paragraphs collection or Style.ParagraphFormat)
'
Public Function Format_SetTabs(ByVal Tabs As Object) As Boolean
Format_SetTabs = False

    If Tabs.TabStops Is Nothing Then Stop: Exit Function
    
    With Tabs.TabStops
    
        .ClearAll
        Dim Inx As Long
        For Inx = 1 To 36
            .Add (Inx * 18)         ' 18 points = .25 inch
        Next Inx
        
    End With

Format_SetTabs = True
End Function

'   Convert a Plain Text Format Item to an HTML Format Item & Cleanup
'
'   !   May create an Inspector for oItem. Cleanup in Caller !
'
Public Function Format_Plain2HTML(ByVal oItem As Object) As Boolean
Format_Plain2HTML = False

    '   If not Plain Text - Done
    '
    If oItem.ItemProperties.Item("BodyFormat") Is Nothing Then
        Format_Plain2HTML = True
        Exit Function
    End If
    If oItem.BodyFormat <> Outlook.olFormatPlain Then
        Format_Plain2HTML = True
        Exit Function
    End If

    '   Convert
    '
    oItem.BodyFormat = Outlook.olFormatHTML
    
    '   Get the InspectorRec and wDoc for the HTML version
    '
    If Not Inspector_ItemInspectorExist(oItem) Then Stop: Exit Function
    Dim InspectorRec As Inspector_InspectorRec
    If Not Inspector_RecFromItem(oItem, InspectorRec) Then Stop: Exit Function
    
    Dim wDoc As Word.Document
    Set wDoc = InspectorRec.oInspector.WordEditor
    If wDoc Is Nothing Then Stop: Exit Function
    If wDoc.ProtectionType <> wdNoProtection Then wDoc.UnProtect
    
    '   Set the Font to Courier
    '
    With wDoc.Content.Font
        .Name = "Courier New"
        .Size = 10
    End With
    
    '   Change all para to single line/paragraph spacing.
    '
    With wDoc.Paragraphs.Format
        .LineSpacingRule = wdLineSpaceSingle
        .SpaceAfterAuto = False
        .SpaceAfter = 0
        .SpaceBeforeAuto = False
        .SpaceBefore = 0
    End With

Format_Plain2HTML = True
End Function

' =====================================================================
'   Plain
' =====================================================================

Public Sub Format_PlainAll(ByVal wRange As Word.Range)

    Dim wPara As Word.Paragraph
    For Each wPara In wRange.Paragraphs
        Format_PlainPara wPara
    Next wPara
    
    Format_PlainFont wRange
    
End Sub

'   Reset, Zero, Default every stupid Para attribute I can find
'
Public Sub Format_PlainPara(ByVal wPara As Word.Paragraph)

    With wPara.Format
    
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
        .PageBreakBefore = False
        .Hyphenation = False
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        
        With .Borders
        
            .DistanceFromTop = 0
            .DistanceFromLeft = 0
            .DistanceFromBottom = 0
            .DistanceFromRight = 0
            .Shadow = False
            
            .Item(wdBorderLeft).LineStyle = wdLineStyleNone
            .Item(wdBorderRight).LineStyle = wdLineStyleNone
            .Item(wdBorderTop).LineStyle = wdLineStyleNone
            .Item(wdBorderBottom).LineStyle = wdLineStyleNone
        
        End With

    End With
    
End Sub

'   Reset, Zero, Default every stupid Font attribute I can find
'
Public Sub Format_PlainFont(ByVal wRange As Word.Range)

    With wRange.Font
    
        .Name = "Courier New"
        .Size = 10
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .Strikethrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorAutomatic
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Spacing = 0
        .Scaling = 100
        .Position = 0
        .Kerning = 0
        .Animation = wdAnimationNone
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    
    End With

End Sub

'   Plain Footer with Text Date/Time on the left and Page on the right
'
Public Sub Format_PlainFooter(ByVal wRange As Word.Range, ByVal Text As String)

    '   2025-02-17 - For complex doc.
    '
    '       Comming in Protected. Even though the Doc was already unprotected.
    '       Works when debugging. So added the DoEvents.
    '       All just guesswork. No idea what Stupid II is doing.
    '
    DoEvents
    If wRange.Document.ProtectionType <> wdNoProtection Then wRange.Document.UnProtect
    DoEvents
    
    '   We now return to our normal programming
    '
    With wRange
    
        '   Clear the footer
        '
        .Delete
        .ParagraphFormat.TabStops.ClearAll
        .Font.Name = "Tahoma"
        .Font.Size = 10
        
        '   Build the footer line BACKWARDS
        '
        Dim FRange As Word.Range
        Set FRange = .Duplicate
        With FRange
        
            .Collapse wdCollapseStart
            .Fields.Add FRange, wdFieldNumPages
            .Collapse wdCollapseStart
            .InsertAfter " of "
            .Collapse wdCollapseStart
            .Fields.Add FRange, wdFieldPage
            .Collapse wdCollapseStart
            .InsertAfter "Page "
            .Collapse wdCollapseStart
            .InsertAlignmentTab wdRight, wdMargin
            .Collapse wdCollapseStart
            .InsertDateTime DateTimeFormat:="HH:mm"
            .Collapse wdCollapseStart
            .InsertAfter " "
            .Collapse wdCollapseStart
            .InsertDateTime DateTimeFormat:="yyyy-MM-dd"
            .Collapse wdCollapseStart
            .InsertAfter " "
            .Collapse wdCollapseStart
            .InsertAfter Text
            
        End With
            
        '   Add a Para Top Border
        '
        With .Paragraphs.Item(1)
            
            With .Borders
            
                .DistanceFromTop = 8
                .DistanceFromLeft = 0
                .DistanceFromBottom = 0
                .DistanceFromRight = 0
                .Shadow = False
                
                .Item(wdBorderLeft).LineStyle = wdLineStyleNone
                .Item(wdBorderRight).LineStyle = wdLineStyleNone
                .Item(wdBorderBottom).LineStyle = wdLineStyleNone
                
                With .Item(wdBorderTop)
                
                    .LineStyle = wdLineStyleSingle
                    .LineWidth = wdLineWidth100pt
                    
                End With
        
            End With
        
        End With
        
    End With

End Sub
