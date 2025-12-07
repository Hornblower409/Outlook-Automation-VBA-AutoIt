Attribute VB_Name = "ProcXeq_"
Option Explicit
Option Private Module

'   Decode and execute a ProcXeq Command Line - Callback from timer started by ProcXeq Instance
'
'       If you change any COMMAND_NAMES you might break an existing ProcXeq external link.
'
Public Function ProcXeq_Decode(ByVal P1 As Long, ByVal P2 As Long, ByVal TimerId As Long, ByVal P3 As Long) As Boolean
Const ThisProc = "ProcXeq_Decode"
ProcXeq_Decode = False

    '   Get my Timer and Disable it
    '
    Dim Timer As Timer: Set Timer = glbAppTimers.Timers(TimerId)
    If Timer Is Nothing Then Stop: Exit Function
    Timer.Disable

    '   Get the Command Line from the Registry
    '   Split it into an Array
    '
    Dim CmdLine() As String
    CmdLine() = Split(glbWshShell.RegRead(glbProcXeq_RegCmdLine), Chr(glbProcXeq_CmdLineSep))
    
    '   Case on the Command (First element of the Command Line Array)
    '   Call out to the macro.
    '   Always fall through so we reset the ProcXeq Status.
    '
    '   SPOS - 2024-10-04
    '
    '       Shutdown and Restart after adding a new Command to the list. Not just App Init.
    '       Or you may get a Status not changed to Running when you try to execute. I have no idea why.
    '
    Select Case UCase(CmdLine(0))
    
        Case "ADDRESSBOOK":                     Mail_AddressBook
        Case "CATEGORIES_ASSIGN":               Categories_Assign
        Case "CATEGORIES_CATSEARCH":            Categories_CatSearch
        Case "CLEANUP_FORWARDNOCLEANUP":        Cleanup_ForwardNoCleanup
        Case "DOWNLOADFOLDER":                  IMAP_DownloadFolder CmdLine(1)
        Case "EDITMODE":                        Ribbon_EditModeActivate ActiveInspector
        Case "EXPORTASPDF":                     Mail_ExportAsPDF
        Case "FOLLOWUP_SET":                    FollowUp_Set
        Case "FOLLOWUP_CLEAR":                  FollowUp_Clear
        Case "FORMAT_BLOCKQUOTE":               Format_BlockQuote
        Case "FORMAT_BULLETTOGGLE":             Format_BulletToggle
        Case "FORMAT_COURIER10":                Format_Courier10
        Case "FORMAT_FONTDIALOG":               Format_FontDialog
        Case "FORMAT_FONTSET":                  Format_FontSet CmdLine()
        Case "FORMAT_GRAY":                     Format_Gray
        Case "FORMAT_NORMAL":                   Format_Normal
        Case "FORMAT_RED":                      Format_Red
        Case "FORMAT_TABS":                     Format_Tabs
        Case "FORWARD":                         Ribbon_ExecuteMSO ActiveWindow, glbidMSO_Forward
        Case "NEWMAIL":                         Mail_NewMailAndAddressBook
        Case "PROJECTS_CATANDMOVE":             Projects_CatAndMove FollowUp:=False
        Case "PROJECTS_CATMOVEANDFOLLOWUP":     Projects_CatAndMove FollowUp:=True
        Case "PROJECTS_CATANDMOVEJUNK":         Projects_CatAndMoveJunk
        Case "PROJECTS_CATANDSEND":             Projects_CatAndSend FollowUp:=False
        Case "PROJECTS_CATSENDANDFOLLOWUP":     Projects_CatAndSend FollowUp:=True
        Case "PROJECTS_CATCHECK":               Projects_CatCheck
        Case "REPLY":                           Ribbon_ExecuteMSO ActiveWindow, glbidMSO_Reply
        Case "REPLYALL":                        Ribbon_ReplyAll
        Case "SENDRECEIVEALL":                  Ribbon_ExecuteMSO ActiveWindow, glbidMSO_SendReceiveAll
        Case "SMARTDEL":                        IMAP_SmartDel
        Case "SMARTDELPURGE":                   IMAP_SmartDelPurge CmdLine(1)
        Case "UPDATEFOLDER":                    IMAP_UpdateFolder CmdLine(1)
        Case "WIPPROJNEW":                      CustForm_WipProjNew
        
        Case "TEST3":                           ProcXeq_TEST3 CmdLine()
        Case Else:                              Stop
    
    End Select
        
    '   Change the ProcXeq Status to Ready
    '
    glbWshShell.RegWrite glbProcXeq_RegStatus, glbProcXeq_Status_Ready, "REG_SZ"
    
ProcXeq_Decode = True
End Function

'   Test Case
'
Public Function ProcXeq_TEST3(ByRef CmdLine() As String) As Boolean
Const ThisProc = "ProcXeq_TEST3"
ProcXeq_TEST3 = False

    If UBound(CmdLine) <> 3 Then
        Msg_Box Proc:=ThisProc, Step:="Check Args", Text:="TEST3 ProcXeq must have exactly 3 args."
        Exit Function
    End If

    Msg_Box Proc:=ThisProc, _
            Text:="I'm a little Test Proc short and stout." & glbBlankLine & _
                   "Arg1 = '" & CmdLine(1) & "'" & vbNewLine & _
                   "Arg2 = '" & CmdLine(2) & "'" & vbNewLine & _
                   "Arg3 = '" & CmdLine(3) & "'"
                   
ProcXeq_TEST3 = True
End Function



