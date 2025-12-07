Attribute VB_Name = "Send_"
Option Explicit
Option Private Module

'   Item Send Event
'
Public Sub Send_ItemSend(ByVal Item As Object, ByRef Cancel As Boolean)

    Cleanup_Send Item, Cancel
    If Cancel Then Exit Sub
    
    Send_MaxSize Item, Cancel
    If Cancel Then Exit Sub
    
    Send_CatAndSend Item, Cancel
    If Cancel Then Exit Sub

End Sub

'   Handle any Cats on an outgoing Item
'
'   If it has Cats and the SaveSentMessageFolder = Projects
'
'       Stash the Cats String (as Hex) in the BillingInformation field.
'       Strip any Cats from the outgoing Item.
'
'   SPOS - Means that things I send that don't have a SaveSentMessageFolder (e.g. Meeting Invites) go out with their Cats intact.
'   SPOS - No way around it. If I strip Cats from a Meeting Invite, the Meeting on my box looses the Cats as well.
'   SPOS - All I can do is warn them when it is sent.
'
Private Sub Send_CatAndSend(ByVal Item As Object, ByRef Cancel As Boolean)
Const ThisProc = "Send_CatAndSend"
Cancel = True

    '   Get the SaveSentFolder (if any)
    '
    Dim SaveSentFolder As Outlook.Folder
    If Not Mail_GetSaveSentFolder(Item, SaveSentFolder) Then Exit Sub
    
    '   Is it going to Projects?
    '
    Dim SaveSentIsProjects As Boolean: SaveSentIsProjects = False
    If Not SaveSentFolder Is Nothing Then SaveSentIsProjects = (SaveSentFolder.FolderPath = glbKnownPath_Projects)
  
    '   If not going to Project
    '
    If Not SaveSentIsProjects Then
    
        '   If has Cats - Warn and get advice
        '
        If Item.Categories <> "" Then

            Select Case Msg_Box( _
                Proc:=ThisProc, Step:="Has Cats Warning", Icon:=vbQuestion, Buttons:=vbYesNo, Default:=vbDefaultButton2, _
                Text:="Item has Categories that will be visable if sent." & glbBlankLine & "Continue to Send?")
            Case vbNo
                Exit Sub
            End Select
            
        End If
        
        Cancel = False
        Exit Sub
    
    End If

    '   Is going to Projects - Must have Cats
    '
    If Item.Categories = "" Then
            Msg_Box Proc:=ThisProc, Step:="Check Item Cats", Text:="Somehow we got an Item with no Cats but SaveSentMessageFolder is Projects."
            Exit Sub
    End If
    
    '   Stash the Cats in the Item BI
    '   Clear the Cats
    '
    If Not Mail_BISet(Item, glbBIInx_Cats, Item.Categories) Then Stop: Exit Sub
    Item.Categories = ""
    
Cancel = False
End Sub

'   Check the total size of the item and warn if over limit
'
'       SPOS - Does not do a max attachments size check on forward/reply
'       even if you add more attachments to the original.
'
'       SPOS - MailItem.Size is Zero unless you Save (to Drafts) first.
'       I don't want to Save really gig ones, so I calculate the size myself
'       from the HTMLBody + attachments
'
Private Sub Send_MaxSize(ByVal Item As Object, ByRef Cancel As Boolean)
Const ThisProc = "Send_MaxSize"

    Dim TotalSize As Long
' Unused 2023-08-10
'    Dim Attachments As Outlook.Attachments
    Dim Attachment As Outlook.Attachment

    '   Body Size
    '
    If Mail_HasHTMLBody(Item) Then
        TotalSize = Len(Item.HTMLBody)
    Else
        TotalSize = Len(Item.Body)
    End If

    '   Plus attachment sizes
    '
    For Each Attachment In Item.Attachments
        TotalSize = TotalSize + Attachment.Size
    Next Attachment
    
    '   Warn if over weight
    '
    If TotalSize < glbMaxMessageSize Then Exit Sub
    
    If Msg_Box( _
        Proc:=ThisProc, Step:="Check total message size", Subject:=Item.Subject, Buttons:=vbYesNo, Default:=vbDefaultButton2, Icon:=vbQuestion, _
        Text:="Anything over " & Misc_BytesToStr(glbMaxMessageSize) & " will cause problems with some email services (e.g. GMail)." & glbBlankLine & _
              "Total message size is " & Misc_BytesToStr(TotalSize) & ". Send it anyway?") _
    <> vbYes Then
        Cancel = True
        Exit Sub
    End If

End Sub
