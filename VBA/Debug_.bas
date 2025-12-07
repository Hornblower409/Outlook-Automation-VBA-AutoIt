Attribute VB_Name = "Debug_"
Option Explicit
Option Private Module

' ---------------------------------------------------------------------
'   Debug.Print Word Range
'
'       All of these are because the VBE blows up if I try to expand a
'       Word Range in the Locals window.
'
'   Call: Debug_PrintRangeXXXXX  wRange, "wRange"
'
' ---------------------------------------------------------------------

'   Debug.Print a Word Range Start and End
'
Public Sub Debug_PrintRangeSpan(ByVal wRange As Word.Range, Optional ByVal RangeName As String = "")

    Debug.Print "---------------------------------------------"
    Debug.Print RangeName, "Span" & vbNewLine
    
    With wRange
    
        Debug.Print ".Start", .Start
        Debug.Print ".End", .End
    
    End With

End Sub

'   Debug.Print a Word Range Text
'
Public Sub Debug_PrintRangeText(ByVal wRange As Word.Range, Optional ByVal RangeName As String = "")

    Debug.Print "============================================="
    Debug.Print RangeName, "Text"
    Debug.Print "---------------------------------------------"
    
    With wRange
    
        Debug.Print .Text
    
    End With
    
    Debug.Print "============================================="
    
End Sub


'   Debug.Print a Word .Find Range
'
Public Sub Debug_PrintRangeFind(ByVal wRange As Word.Range, Optional ByVal RangeName As String = "")

    Debug.Print "---------------------------------------------"
    Debug.Print RangeName, ".Find" & vbNewLine
    
    With wRange
    
        With .Find
        
            Debug.Print ".Foramt", .Format
            Debug.Print ".Forward", .Forward
            Debug.Print ".Found", .Found
            Debug.Print ".IgnorePunct", .IgnorePunct
            Debug.Print ".IgnoreSpace", .IgnoreSpace
            Debug.Print ".MatchAllWordForms", .MatchAllWordForms
            Debug.Print ".MatchCase", .MatchCase
            Debug.Print ".MatchControl", .MatchControl
            Debug.Print ".MatchPrefix", .MatchPrefix
            Debug.Print ".MatchSoundsLike", .MatchSoundsLike
            Debug.Print ".MatchSuffix", .MatchSuffix
            Debug.Print ".MatchWholeWord", .MatchWholeWord
            Debug.Print ".MatchWildcards", .MatchWildcards
            Debug.Print ".Text", .Text
            Debug.Print ".Wrap", IIf(.Wrap = wdFindAsk, "Ask", IIf(.Wrap = wdFindContinue, "Continue", IIf(.Wrap = wdFindStop, "Stop", "!Oops")))
            
        End With
        
    End With
    
End Sub

'   Debug.Print a Word .ParagraphFormat Range
'
Public Sub Debug_PrintRangeFormat(ByVal wRange As Word.Range, Optional ByVal RangeName As String = "")

    Debug.Print "---------------------------------------------"
    Debug.Print RangeName, ".ParagraphFormat" & vbNewLine

    With wRange

            With .ParagraphFormat
            
                Debug.Print ".AddSpaceBetweenFarEastAndAlpha", .AddSpaceBetweenFarEastAndAlpha
                Debug.Print ".AddSpaceBetweenFarEastAndDigit", .AddSpaceBetweenFarEastAndDigit
                Debug.Print ".Alignment", .Alignment
                Debug.Print ".Application", .Application.Name
                Debug.Print ".AutoAdjustRightIndent", .AutoAdjustRightIndent
                Debug.Print ".BaseLineAlignment", .BaseLineAlignment
                Debug.Print ".CharacterUnitFirstLineIndent", .CharacterUnitFirstLineIndent
                Debug.Print ".CharacterUnitLeftIndent", .CharacterUnitLeftIndent
                Debug.Print ".CharacterUnitRightIndent", .CharacterUnitRightIndent
                Debug.Print ".Creator", .Creator
                Debug.Print ".DisableLineHeightGrid", .DisableLineHeightGrid
                Debug.Print ".FarEastLineBreakControl", .FarEastLineBreakControl
                Debug.Print ".HalfWidthPunctuationOnTopOfLine", .HalfWidthPunctuationOnTopOfLine
                Debug.Print ".HangingPunctuation", .HangingPunctuation
                Debug.Print ".Hyphenation", .Hyphenation
                Debug.Print ".KeepTogether", .KeepTogether
                Debug.Print ".KeepWithNext", .KeepWithNext
                Debug.Print ".LeftIndent", .LeftIndent
                Debug.Print ".LineSpacing", .LineSpacing
                Debug.Print ".LineSpacingRule", .LineSpacingRule
                Debug.Print ".LineUnitAfter", .LineUnitAfter
                Debug.Print ".LineUnitBefore", .LineUnitBefore
                Debug.Print ".MirrorIndents", .MirrorIndents
                Debug.Print ".NoLineNumber", .NoLineNumber
                Debug.Print ".OutlineLevel", .OutlineLevel
                Debug.Print ".PageBreakBefore", .PageBreakBefore
                Debug.Print ".ReadingOrder", .ReadingOrder
                Debug.Print ".RightIndent", .RightIndent
                Debug.Print ".SpaceAfter", .SpaceAfter
                Debug.Print ".SpaceAfterAuto", .SpaceAfterAuto
                Debug.Print ".SpaceBefore", .SpaceBefore
                Debug.Print ".SpaceBeforeAuto", .SpaceBeforeAuto
                
                ' Debug.Print ".Style", .Style  --> BOOM
                Debug.Print ".Style", "Can't touch this"
                
                Debug.Print ".TextboxTightWrap", .TextboxTightWrap
                Debug.Print ".WidowControl", .WidowControl
                Debug.Print ".WordWrap", .WordWrap
            
            End With
        
        End With

End Sub

' =====================================================================
'   HotRod Log File
' =====================================================================

'   Write a line to the HotRod Log File
'
Public Function Debug_HotRodLog( _
    Optional ByVal Proc As String = "", _
    Optional ByVal Step As String = "", _
    Optional ByVal oErr As VBA.ErrObject = Nothing, _
    Optional ByVal Subject As String = "", _
    Optional ByVal Text As String = "", _
    Optional ByVal P1Name As String = "", _
    Optional ByVal P1Value As String = "", _
    Optional ByVal P2Name As String = "", _
    Optional ByVal P2Value As String = "", _
    Optional ByVal P3Name As String = "", _
    Optional ByVal P3Value As String = "", _
    Optional ByVal P4Name As String = "", _
    Optional ByVal P4Value As String = "" _
) As Boolean
Const ThisProc = "Debug_HotRodLog"
Debug_HotRodLog = False

    '   Get any oErr values before I clear it with "On Error GoTo 0"
    '
    If Not (oErr Is Nothing) Then
        Dim ErrDec As String: ErrDec = oErr.Number
        Dim ErrHex As String: ErrHex = "0x" & Hex(Err.Number)
        Dim ErrDesc As String: ErrDesc = oErr.Description
    End If
    On Error GoTo 0
    
    '   Define the Columns and Row
    '
    Const colFirst      As Long = 0
    Const colStamp      As Long = 0
    Const colProc       As Long = 1
    Const colStep       As Long = 2
    Const colText       As Long = 3
    Const colSubject    As Long = 4
    Const colErrDec     As Long = 5
    Const colErrHex     As Long = 6
    Const colErrDesc    As Long = 7
    Const colP1Name     As Long = 8
    Const colP1Value    As Long = 9
    Const colP2Name     As Long = 10
    Const colP2Value    As Long = 11
    Const colP3Name     As Long = 12
    Const colP3Value    As Long = 13
    Const colP4Name     As Long = 14
    Const colP4Value    As Long = 15
    Const colLast       As Long = 15
    
    Dim aRow() As String
    
    '   If Stupid cleared all my Globals
    '
    If glbFSO Is Nothing Then Set glbFSO = New Scripting.FileSystemObject
    
    '   If there is no Log file - write the Header Row
    '
    If Not glbFSO.FileExists(glbHotRodLogFile) Then
    
        ReDim aRow(colFirst To colLast)
        
        aRow(colStamp) = "Time"
        aRow(colProc) = "Proc"
        aRow(colStep) = "Step"
        aRow(colErrDec) = "ErrDec"
        aRow(colErrHex) = "ErrHex"
        aRow(colErrDesc) = "ErrDesc"
        aRow(colSubject) = "Subject"
        aRow(colText) = "Text"
        aRow(colP1Name) = "P1 Name"
        aRow(colP1Value) = "P1 Value"
        aRow(colP2Name) = "P2 Name"
        aRow(colP2Value) = "P2 Value"
        aRow(colP3Name) = "P3 Name"
        aRow(colP3Value) = "P3 Value"
        aRow(colP4Name) = "P4 Name"
        aRow(colP4Value) = "P4 Value"
        
        If Not File_AppendText(glbHotRodLogFile, Join(aRow, vbTab)) Then Stop: Exit Function
        
    End If

    '   Build and write the Row
    '
    ReDim aRow(colFirst To colLast)
    
    aRow(colStamp) = Misc_NowStamp()
    aRow(colProc) = Proc
    aRow(colStep) = Step
    aRow(colErrDec) = ErrDec
    aRow(colErrHex) = ErrHex
    aRow(colErrDesc) = ErrDesc
    aRow(colSubject) = Subject
    aRow(colText) = Text
    aRow(colP1Name) = P1Name
    aRow(colP1Value) = P1Value
    aRow(colP2Name) = P2Name
    aRow(colP2Value) = P2Value
    aRow(colP3Name) = P3Name
    aRow(colP3Value) = P3Value
    aRow(colP4Name) = P4Name
    aRow(colP4Value) = P4Value

    If Not File_AppendText(glbHotRodLogFile, Join(aRow, vbTab)) Then Stop: Exit Function

Debug_HotRodLog = True
End Function
