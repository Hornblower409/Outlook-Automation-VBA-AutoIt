VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcXeq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents JEvents As Outlook.JournalItem
Attribute JEvents.VB_VarHelpID = -1

Private Const RSVPJournalItemName As String = "ProcXeq Journal Entry For Event Trigger by External Program"

Private Sub JEvents_CustomAction(ByVal Action As Object, ByVal Response As Object, ByRef Cancel As Boolean)
Const ThisProc = "ProcXeq.JEvents_CustomAction"

    '   If the Action is not my Custom Action - Whoa!
    '
    If Action.Name <> glbProcXeq_ActionName Then Stop: Exit Sub
    
    '   Ditch the Response
    '   Setup to Cancel the Action on return
    '
    If Not Response Is Nothing Then Response.Close Outlook.OlInspectorClose.olDiscard
    Cancel = True
    
    '   Case on the current Registry Status
    '
    Dim Status As String
    Status = glbWshShell.RegRead(glbProcXeq_RegStatus)
    Select Case Status
    
        '   Canceled - AutoIt ProcXeq timed out waiting for me
        '
        Case glbProcXeq_Status_Canceled
        
            '   ProcXeq Status <- Ready and done
            '
            glbWshShell.RegWrite glbProcXeq_RegStatus, glbProcXeq_Status_Ready, "REG_SZ"
            Exit Sub

        '   Submitted - Continue
        '
        Case glbProcXeq_Status_Submitted
            
        '   Else - Error
        '
        Case Else
        
            Msg_Box Proc:=ThisProc, Step:="Check Status", _
                    Text:="ProcXeq current Status is '" & Status & "'." & vbNewLine & _
                          "Must be '" & glbProcXeq_Status_Submitted & "'."
            Exit Sub
            
    End Select
    
    '   Change ProcXeq Status to Running
    '   Start the timer to callback at ProcXeq_Decode
    '
    glbWshShell.RegWrite glbProcXeq_RegStatus, glbProcXeq_Status_Running, "REG_SZ"
    If Not Decode() Then Stop: Exit Sub
    
End Sub

'   Use a timer to call ProcXeq_Decode to give me time to Cancel the Event and Return
'
Private Function Decode() As Boolean
Const ThisProc = "ProcXeq.Decode"
Decode = False

    '   Build my Timer
    '
    Dim Timer As Timer: Set Timer = New Timer
    With Timer
    
        .Name = ThisProc
        .Period = 50
        .Retries = 0
        .Callback AddressOf ProcXeq_Decode
        
    End With
    
    '   Start my Timer
    '
    glbAppTimers.Add Timer
    Timer.Enable

Decode = True
End Function

'   Instance Initialize
'
Public Function Initialize() As Boolean
Const ThisProc = "ProcXeq.Initialize"
Initialize = False

    '   Get the Journal Folder
    '
    Dim JFolder As Outlook.Folder
    Set JFolder = Folders_KnownPath(glbKnownPath_Journal)
    
    '   Look for my RSVP Journal Item
    '
    Dim JItem As Outlook.JournalItem
    Dim JItemFound As Boolean: JItemFound = False
    For Each JItem In JFolder.Items
        If JItem.Subject = RSVPJournalItemName Then
            JItemFound = True
            Exit For
        End If
    Next JItem

    '   If no RSVP Journal Item - create a new one
    '
    If Not JItemFound Then If Not NewJItem(JItem) Then Stop: Exit Function
    
    '   Hook the events on the RSVP Journal Item
    '
    Set JEvents = JItem
    
    '   Write the RSVP Journal Item Entry/Store IDs to the registry
    '
    glbWshShell.RegWrite glbProcXeq_RegBaseKey, ""
    glbWshShell.RegWrite glbProcXeq_RegEntryId, JItem.EntryId, "REG_SZ"
    glbWshShell.RegWrite glbProcXeq_RegStoreId, JFolder.StoreID, "REG_SZ"
    
    '   Change the ProcXeq Status to Ready
    '
    glbWshShell.RegWrite glbProcXeq_RegStatus, glbProcXeq_Status_Ready, "REG_SZ"

Initialize = True
End Function

'   Create a new RSVP Journal Item
'
Private Function NewJItem(ByRef JItem As Outlook.JournalItem) As Boolean
Const ThisProc = "NewJItem"
NewJItem = False

    Dim JFolder As Outlook.Folder
    Set JFolder = Folders_KnownPath(glbKnownPath_Journal)
    
    Set JItem = Application.CreateItem(Outlook.olJournalItem)
    With JItem
        .Subject = RSVPJournalItemName
        .Type = RSVPJournalItemName
    End With

    Dim JAction As Outlook.Action
    Set JAction = JItem.Actions.Add
    With JAction
        .Name = glbProcXeq_ActionName
        .Enabled = True
        .MessageClass = "IPM.Note"
        .CopyLike = Outlook.olForward
        .Prefix = ""
        .ReplyStyle = Outlook.olOmitOriginalText
        .ResponseStyle = Outlook.olPrompt
        
        '   SPOS - Maybe. I could be missing something. No matter what,
        '   it does not show the Custom Action anywhere on the Explorer
        '   Context Menu or the Inspector.
        '
        .ShowOn = Outlook.OlActionShowOn.olMenuAndToolbar
        
    End With

    JItem.Save

NewJItem = True
End Function

