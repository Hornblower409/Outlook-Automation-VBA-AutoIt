Attribute VB_Name = "Word_"
Option Explicit
Option Private Module

'   Reset a Word .Find object to defaults
'
'   From https://gregmaxey.com/word_tip_pages/words_fickle_vba_find_property.html
'
Public Function Word_FindDefault(ByVal wRange As Word.Range) As Word.Range

    Set Word_FindDefault = wRange
    With Word_FindDefault.Find
    
        .ClearFormatting
        .Format = False
        .Forward = True
        .Highlight = wdUndefined
        .IgnorePunct = False
        .IgnoreSpace = False
        .MatchAllWordForms = False
        .MatchCase = False
        .MatchPhrase = False
        .MatchPrefix = False
        .MatchSoundsLike = False
        .MatchSuffix = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .Replacement.ClearFormatting
        .Replacement.Text = ""
        .Text = ""
        .Wrap = wdFindStop

    End With
    
End Function

'   Get the Paragraph Index of a Paragraph in a Range
'
'       Returns FALSE if the Para is not in the Range
'
Public Function Word_ParaIndex(ByVal wPara As Word.Paragraph, ByVal wRange As Word.Range, ByRef ParaIndex As Long) As Boolean

    Word_ParaIndex = False

    If Not wPara.Range.InRange(wRange) Then Exit Function
    
    '   wRangeSpan = Everything from wRange.Start to (and including) wPara
    '
    Dim wRangeSpan As Word.Range
    Set wRangeSpan = wRange.Duplicate
    wRangeSpan.End = wPara.Range.End
    
    '   wPara is the last paragraph in wRangeSpan
    '
    ParaIndex = wRangeSpan.Paragraphs.count
    
    Word_ParaIndex = True

End Function

'   Do a Word Text Find/Replace on a Range. Returns TRUE/FALSE.
'
'   Word_Replace( Range, Find, Replace, {wdReplaceAll, wdReplaceOne} )
'
'   SPOS -  The ^u notation will not work as the Replacement Text. Use CharW(nnnn).
'
Public Function Word_Replace(ByVal wRange As Word.Range, ByVal sFind As String, ByVal sReplace As String, ByVal wReplace As Word.WdReplace) As Boolean

    Word_Replace = False
    
    Dim wSearch As Word.Range
    Set wSearch = Word_FindDefault(wRange.Duplicate)
    wSearch.Find.Text = sFind
    
    wSearch.Find.Replacement.Text = sReplace
    Word_Replace = wSearch.Find.Execute(Replace:=wReplace)

End Function

'   Do a Word Text Find on a Range and keep the right N chars of the Found string. Returns TRUE/FALSE
'
'   Word_DeleteLeft( Range, Find, Count, {wdReplaceAll, wdReplaceOne} )
'
'       SPOS - Some ^p can not be replaced by Word Find/Replace.
'
'       Find.Execute is TRUE but target is still there. Same in Word GUI. It finds it but
'       won't replace it. Is NOT a "<w:cr/>" in the XML that I can see.
'
'       Typically it's a naked ^p just above a table or the "End Of Doc ^p"
'
'       This proc let's me do things to what is in front of a stuck ^p.
'
Public Function Word_DeleteLeft(ByVal wRange As Word.Range, ByVal sFind As String, ByVal RCount As Long, ByVal wReplace As Word.WdReplace) As Boolean

    Word_DeleteLeft = False
    
    Dim wSearch As Word.Range
    Set wSearch = Word_FindDefault(wRange.Duplicate)
    wSearch.Find.Text = sFind
    
    Dim wDel As Word.Range
    Set wDel = wRange.Duplicate
    
    Do
    
        '   Reset the Search Range
        '
        wSearch.Start = wRange.Start
        wSearch.End = wRange.End
    
    '   While Find is True
    '
    If Not wSearch.Find.Execute Then Exit Do
    
        '   If we had at least one hit - set return TRUE
        '
        Word_DeleteLeft = True
    
        '   Delete anything left of RCount from the end
        '
        wDel.Start = wSearch.Start
        wDel.End = wSearch.End - RCount
        wDel.Delete
        
        '   If a one off - done
        '
        If wReplace = wdReplaceOne Then Exit Do
        
    Loop

End Function

'   Do a Word Text Find on a Range. Returns TRUE/FALSE.
'
'   Word_Find( Range, Find )
'
Public Function Word_Find(ByVal wRange As Word.Range, ByVal sFind As String) As Boolean

    Dim wSearch As Word.Range
    Set wSearch = Word_FindDefault(wRange.Duplicate)
    
    wSearch.Find.Text = sFind
    Word_Find = wSearch.Find.Execute

End Function

'   Do a Word Text Find on a Range. Returns the Found Range or Nothing.
'
'   Word_FindRange( Range, Find )
'
Public Function Word_FindRange(ByVal wRange As Word.Range, ByVal sFind As String, Optional ByVal Backwards As Boolean = False) As Word.Range

    Set Word_FindRange = Word_FindDefault(wRange.Duplicate)
    
    With Word_FindRange.Find
    
        .Forward = Not Backwards
        .Text = sFind
        
        If Not .Execute Then Set Word_FindRange = Nothing
    
    End With
    
End Function

'   Do a Word Text Find on a Range that acts like an InStr()
'
'   Word_FindInStrRange( Range, Find )
'   Returns Found.Start/End = 0 if not found.
'
'   !  Watch out if you mix this Find with my others (or MS Standard)  !
'   !  You will need to adjust the Found.Start = Found.Start - 1       !
'
'   SPOS - The way the Word Find works screws up my brain. The Found.Start is one char
'   BEFORE the actual char position in the range. (e.g. If the found string
'   is the first thing in a range then Found.Start = 0). This is an attempt to make
'   a Find that acts like an InStr() function call.
'
'   If the Range.Start > 0 then back it up one (so it won't skip the first char in the Range).
'   If the .Execute was TRUE then Found.Start = Found.Start + 1 (so it points to the first char)
'
Public Function Word_FindInStrRange(ByVal wRange As Word.Range, ByVal sFind As String) As Word.Range

    Set Word_FindInStrRange = wRange.Duplicate
    If Word_FindInStrRange.Start > 0 Then Word_FindInStrRange.Start = Word_FindInStrRange.Start - 1
    Set Word_FindInStrRange = Word_FindDefault(Word_FindInStrRange)
    
    With Word_FindInStrRange
    
        .Find.Text = sFind
        If Not .Find.Execute Then
            .Start = 0
            .End = 0
        Else
            .Start = .Start + 1
        End If
    
    End With
    
End Function

'   Find a Word Para Style (Name or Object) in a Range. Returns the Found Range.
'
Public Function Word_FindParaStyleRange(ByVal wRange As Word.Range, ByVal style As Variant) As Word.Range

    Set Word_FindParaStyleRange = Word_FindDefault(wRange.Duplicate)

    '   If Style does not exsist in this Doc - done
    '
    If Not Word_StyleExists(wRange, style) Then GoTo ErrExit
    
    '   Search
    '
    With Word_FindParaStyleRange.Find
    
        .Format = True
        
        '   SPOS - You can NOT do Find.ParagraphFormat.Style = Style. It barfs.
        '   You have to give it a WHOLE ParagraphFormat. And if you do give it the
        '   whole thing (Find.ParagraphFormat = wStyle.ParagraphFormat) it won't find anything.
        '
        '   So (as far as I can tell) there is no way to find a Style that is being used ONLY
        '   as a Paragraph Style, without checking the results. Since using Find.Style will
        '   find the Style when used as a Character Style as well (for dual use Styles).
        '
        .style = style
        If .Execute Then Exit Function
        
    End With
    
ErrExit:

    Word_FindParaStyleRange.Start = 0
    Word_FindParaStyleRange.End = 0
    
End Function

'   Does a Style (Name or Object) exist in a Range?
'
Public Function Word_StyleExists(ByVal wRange As Word.Range, ByVal style As Variant) As Boolean

    Word_StyleExists = False

    '   SPOS - The only way to test for existance is an error trap.
    '   See https://roxtonlabs.blogspot.com/2015/09/vba-test-if-style-exists-in-word.html
    '
    '   Existence is in the Doc (Range.Parent) actually, but I use ranges everywhere, so this makes it easier.
    '
    Dim wStyle As Word.style
    On Error Resume Next
    
        Set wStyle = wRange.Parent.Styles(style)
        If Err.Number = glbError_MemberNotFound Then Exit Function
        If Err.Number <> glbError_None Then Stop: Exit Function
        
    On Error GoTo 0
        
    Word_StyleExists = True
    
End Function

