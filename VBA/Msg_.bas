Attribute VB_Name = "Msg_"
Option Explicit
Option Private Module

' ---------------------------------------------------------------------
'   Constants
'
'       Icon:=
'
'           vbCritical
'           vbQuestion
'           vbExclamation
'           vbInformation
'
' ---------------------------------------------------------------------
'   Typical Usage:
'
'   Just a message
'
'       Msg_Box Proc:=ThisProc, Text:="Message Text"
'       Msg_Box Proc:=ThisProc, Text:="Message Text", Subject:=SrcItem.Subject
'       Msg_Box Proc:=ThisProc, Text:="Message Text", Step:="Step Name"
'       Msg_Box Proc:=ThisProc, Step:="Step Name", Text:="Message Text"
'       Msg_Box Proc:=ThisProc, Step:="Step Name", Text:="Message Text", Subject:=SrcItem.Subject
'
'   System Error
'
'       Msg_Box oErr:=Err, Proc:=ThisProc, Step:="Step Name", Subject:=Item.Subject
'       Msg_Box oErr:=Err, Proc:=ThisProc, Step:="Step Name", Text:="Message Text", Subject:=Item.Subject
'
'   Trap and Choice
'
'       On Error Resume Next
'
'           {Thing that night cause an error}
'
'           Select Case Err.Number
'               Case glbError_None, Error01, ..., ErrorNN
'                   ' Continue
'               Case Else
'                   Stop: Exit Sub
'           End Select
'
'       On Error GoTo 0
'
'   Question
'
'    Select Case Msg_Box(Text:="Message Text." & glbBlankLine & "Question?", Icon:=vbQuestion, Buttons:=vbYesNoCancel, Default:=vbDefaultButton2, Proc:=ThisProc, Step:="Step Name")
'        Case vbYes
'        Case vbNo
'        Case vbCancel
'    End Select
'
'    Select Case Msg_Box( _
'            Proc:=ThisProc, Step:="Step Name", _
'            Icon:=vbQuestion, Buttons:=vbYesNoCancel, Default:=vbDefaultButton2, _
'            Subject:="Subject", _
'            Text:="Message Text." & glbBlankLine & _
'            "Question?")
'        Case vbYes
'        Case vbNo
'        Case vbCancel
'    End Select
'
' ---------------------------------------------------------------------

Public Function Msg_Box( _
    Optional ByVal oErr As VBA.ErrObject = Nothing, _
    Optional ByVal Proc As String = "", _
    Optional ByVal Step As String = "", _
    Optional ByVal Text As String = "", _
    Optional ByVal Subject As String = "", _
    Optional ByVal Buttons As Long = vbOKOnly, _
    Optional ByVal Default As Long = vbDefaultButton1, _
    Optional ByVal Icon As Long = vbExclamation _
    ) As Integer
    
    '   Make the Title long so the dialog will be wide.
    '   Looks like 108 is max (depends on character widths).
    '   After that the title ends in "..."
    '
    Dim Title As String
    Title = "Outlook VBA Code" & Space(92)
    
    Dim ProcLine As String: ProcLine = ""
    If Proc <> "" Then ProcLine = vbNewLine & "Proc: '" & Proc & "'."
    
    Dim StepLine As String: StepLine = ""
    If Step <> "" Then StepLine = vbNewLine & "Step: '" & Step & "'."
    
    Dim SubjectLine As String: SubjectLine = ""
    If Subject <> "" Then SubjectLine = glbBlankLine & "Subject: '" & Subject & "'."
    
    Dim ErrText As String: ErrText = ""
    If Not oErr Is Nothing Then
        ErrText = _
        glbBlankLine & "Err.Number: " & Err.Number & " (0x" & Hex(Err.Number) & ")." & _
        glbBlankLine & "Error.Description: '" & Err.Description & "'"
    End If
    
    Dim TextLine As String: TextLine = ""
    If Text <> "" Then TextLine = glbBlankLine & Text
    
    '   Build final MsgBlock and Remove any Leading/Trailing vbNewLine
    '
    Dim MsgBlock As String
    MsgBlock = ProcLine & StepLine & SubjectLine & ErrText & TextLine
    
    While Left(MsgBlock, 2) = vbNewLine
        MsgBlock = Mid(MsgBlock, 3)
    Wend
    While Right(MsgBlock, 2) = vbNewLine
        MsgBlock = Mid(MsgBlock, 1, Len(MsgBlock) - 2)
    Wend
    
    Msg_Box = MsgBox( _
        MsgBlock, _
        Buttons + Default + Icon, _
        Title _
    )

End Function

