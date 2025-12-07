Attribute VB_Name = "Response_"
Option Explicit
Option Private Module

'   Handle all Mail Responses - Called from Reply, ReplyAll, and Forward event trigges
'
'   2024-09-14 - SPOS. Is NOT called for Meeting Accept, Tenative, Decline or Meeting Respond Reply, Reply All, Forward.
'   SPOS thinks they are NEW Items. i.e. Len(Item.EntryID) = 0.
'
'   SPOS - Response Body Text and HTML not synced to the Doc
'
'       In a Response or Open event handler the Response Item Word Doc is fully formed but the
'       Plain Text Body will be the original message and the HTML <body> will be empty. These fields
'       are not updated from the Word Doc until the first Activate (i.e. there was a .Display
'       either auto or from my code).
'
'       Brian says "Somebody added the GET [for Body and HTML] after the properties existed,
'       probably because they needed it for a specific thing and would "absolutely come back
'       and clean it up later"
'
Public Sub Response_Main(ByVal Original As Object, ByRef Response As Object, ByRef Cancel As Boolean, ByVal ResponseType As Integer)

    '   SPOS (Sort Of) - Multiple Response Events.
    '
    '       Events will fire for EVERY Inspector showing the item and EVERY Explorer where
    '       it is the Current Selection. Makes sense, but spooky.
    '
    '       Tried using Response.ConversationIndex to keep from processing the same response multiple times.
    '       Works fine for Mail. The ConversationIndex is the same on every event. But Stupid uses different
    '       values for my Custom Forms for every event. (Why does everybody hate my Custom forms?)
    '
    '       So now I just hold on to the Response object. I thought Stupid would get pissed about me
    '       holding a reference for so long, but he just invalidates the reference when the Response object
    '       gets sent, discarded, etc. (It can't be this easy. Something will blow up.)
    '
    '
    If glbLastResponseObject Is Response Then Exit Sub
    Set glbLastResponseObject = Response
        
    '   If Original is RTF - bail out
    '
    If Mail_IsRTF(Original) Then Exit Sub
    
    '   If Original is a Meeting Item - Special
    '
    If Mail_IsMeetingResponse(Original) Or Mail_IsMeetingRequest(Original) Then
        If Not Response_MeetingItem(Original, Cancel, ResponseType) Then Stop: Exit Sub
        Exit Sub
    End If
    
    '   Do Steps. If any step crapped out - Cancel
    '
    If Not Response_Steps(Original, Response, Cancel, ResponseType) Then
    
        Cancel = True
        Exit Sub
        
    End If
    
End Sub

'   Response - Steps
'
Private Function Response_Steps(ByVal Original As Object, ByRef Response As Object, ByRef Cancel As Boolean, ByVal ResponseType As Integer) As Boolean

    Response_Steps = False

    If Not Response_Plain2HTML(Original, Response, ResponseType) Then Exit Function
    If Not Response_ForceReply(Original, Response) Then Exit Function
    If Not CustForm_Response(Original, Response, ResponseType) Then Exit Function
    Categories_RemoveSpecialCats Response
    If Not Response_All(Original, Response, ResponseType) Then Exit Function
    If Not Response_Attachments(Original, Response, ResponseType) Then Exit Function
    If Not Response_PlainShow(Original, Response, Cancel) Then Exit Function
    
    Response_Steps = True
    
End Function

'   Response - Plain Text - Create a new HTML Response from a Plain Text Original
'
Private Function Response_Plain2HTML(ByVal Original As Object, ByRef Response As Object, ByVal ResponseType As Integer) As Boolean
Response_Plain2HTML = False

    '   If the Original was not Plain Text - done
    '
    If Original.BodyFormat <> Outlook.olFormatPlain Then
        Response_Plain2HTML = True
        Exit Function
    End If
    
    '   Create a new HTML Response
    '
    If Not Response_PlainClone(Original, Response, ResponseType) Then Exit Function
    
    '   Set the Font to Courier
    '
    Dim wDoc As Word.Document
    Set wDoc = Response.GetInspector.WordEditor
    
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
    
Response_Plain2HTML = True
End Function

'   Response - Plain Text - Show the new HTML Response and Cancel the current Plain Text Response
'
Private Function Response_PlainShow(ByVal Original As Object, ByRef Response As Object, ByRef Cancel As Boolean) As Boolean

    Response_PlainShow = True

    '   If the Original Item was not Plain Text - done
    '
    If Original.BodyFormat <> Outlook.olFormatPlain Then Exit Function
    
    '   Show the HTML Response
    '   Cancel the Original Plain Text Response
    '
    Response.GetInspector.Activate
    Cancel = True
        
End Function

'   Response - Plain Text - Replace the current Plain Text Response item with a new HTML Response item
'
Private Function Response_PlainClone(ByVal Original As Object, ByRef Response As Object, ByVal ResponseType As Integer) As Boolean
Response_PlainClone = False

    '   Clone the Original
    '
    Dim Copy As Object
    If Not File_CloneInDeleted(Original, Copy) Then Exit Function
    
    '   Change the Original Format to HTML
    '
    Copy.BodyFormat = Outlook.olFormatHTML
    
    '   Create a new Response from the HTML Original
    '   (Type = whatever the caller ask for)
    '
    Select Case ResponseType
        Case glbResponse_Reply:         Set Response = Copy.Reply
        Case glbResponse_ReplyAll:      Set Response = Copy.ReplyAll
        Case glbResponse_Forward:       Set Response = Copy.Forward
        Case Else
            Stop: Exit Function
    End Select

    '   Delete the Copy
    '
    If Not File_ItemDelete(Copy) Then Stop: Exit Function
    
Response_PlainClone = True
End Function

'   Response - Set InternetReplyID if not already set
'
'       This gives me a sure fire way to know if something is a Response.
'       (Stupid doesn't always add this Property or set it to a value).
'
Private Function Response_ForceReply(ByVal Original As Object, ByVal Response As Object) As Boolean
Const ThisProc = "Response_ForceReply"

    Response_ForceReply = False

    '   If Response InternetReplyID already set - done
    '
    If Mail_IsResponse(Response) Then Response_ForceReply = True: Exit Function

    '   If not set (happens on Post Forward and sometimes on Incomming)
    '   Get the InternetMsgID from the Original
    '   If none - use my GUID string.
    '
    Dim InternetMsgID As String
    InternetMsgID = IMAP_MsgID(Original)
    If InternetMsgID = "" Then InternetMsgID = IMAP_NewInternetMsgID()
    
    '   Set Response InternetReplyID = Original InternetMsgID
    '
    If Not Misc_OLSetProperty(Response, glbPropTag_InternetReplyID, InternetMsgID) Then Exit Function
    
    Response_ForceReply = True
    
End Function

'   Response - Fixup the CC: on a Reply All
'
Private Function Response_All(ByVal Original As Object, ByVal Response As Object, ByVal ResponseType As Integer) As Boolean

    Response_All = False

    If ResponseType <> glbResponse_ReplyAll Then Response_All = True: Exit Function

    '   Walk the Response Recipients
    '
    Dim NoTo As Boolean
    NoTo = True
        
    Dim ResponseRecipient As Outlook.Recipient
    For Each ResponseRecipient In Response.Recipients
    
        '   The Original Sender becomes To:
        '   All else become CC:
        '
        If ResponseRecipient.Address = Original.Sender.Address Then
            ResponseRecipient.Type = Outlook.olTo
            NoTo = False
        Else
            ResponseRecipient.Type = Outlook.olCC
        End If
        
    Next
    
    '   If only one Recipient - make them To:
    '
    If Response.Recipients.count = 1 Then
        Response.Recipients(1).Type = Outlook.olTo
        NoTo = False
    End If
    
    '   If No To: - use original To:
    '
    If NoTo Then
    
        For Each ResponseRecipient In Response.Recipients
        
            Dim OriginalRecipient As Outlook.Recipient
            For Each OriginalRecipient In Original.Recipients
            
                If ResponseRecipient.Address = OriginalRecipient.Address Then
                    If OriginalRecipient.Type = Outlook.olTo Then
                        ResponseRecipient.Type = Outlook.olTo
                        NoTo = False
                    End If
                End If
            
            Next
            
        Next

    End If
    
    '   If No To: - make them all To:
    '   (Response to an email I sent originally)
    '
    If NoTo Then
    
        For Each ResponseRecipient In Response.Recipients
            ResponseRecipient.Type = Outlook.olTo
            NoTo = False
        Next
        
    End If
    
    Response_All = True
        
End Function

'   Response - Copy attachments from the Original to the Response
'
Private Function Response_Attachments(ByVal Original As Object, ByVal Response As Object, ByVal ResponseType As Integer) As Boolean
Const ThisProc = "Response_Attachments"
Response_Attachments = False
    
    '   If no attachments - Done
    '   If not a Mail item (e.g. Post, Note, ...) - Done
    '
    If Original.Attachments.count = 0 Then Response_Attachments = True: Exit Function
    If Original.Class <> Outlook.olMail Then Response_Attachments = True: Exit Function
    
    '   Ask what to do?
    '
    Dim IncludeAttachments As Boolean
    Select Case Msg_Box( _
            Proc:=ThisProc, Step:="User Choice", _
            Icon:=vbQuestion, Buttons:=vbYesNoCancel, Default:=vbDefaultButton2, _
            Subject:=Original.Subject, _
            Text:="Include attachments from Original in Response?")
        Case vbYes
            IncludeAttachments = True
        Case vbNo
            IncludeAttachments = False
        Case vbCancel
            Exit Function
    End Select
    
    '   Do it based on Response Type
    '
    Dim Attachment As Outlook.Attachment
    Select Case ResponseType
    
        '   Forward has the attachments by default. Remove them if asked.
        '
        Case glbResponse_Forward
        
            If IncludeAttachments Then Response_Attachments = True: Exit Function
            
            '   Attachments is a Collection. Can't use a "Forward For" when deleting entries.
            '
            While (Response.Attachments.count > 0): Response.Attachments(1).Delete: Wend
        
        '   Reply/ReplyAll do not have the attachments. Add them if asked.
        '
        Case glbResponse_Reply, glbResponse_ReplyAll
        
            If Not IncludeAttachments Then Response_Attachments = True: Exit Function
            For Each Attachment In Original.Attachments
            
                Dim FileSpec As String
                FileSpec = glbTempFilePath & File_CleanupNameSegment(Attachment.FileName)
                Attachment.SaveAsFile FileSpec
                
                On Error Resume Next
                    Response.Attachments.Add FileSpec, Attachment.Type
                    If Err.Number <> glbError_None Then
                        Msg_Box oErr:=Err, Proc:=ThisProc, Step:="Add Attachments", Subject:=Original.Subject
                        glbFSO.DeleteFile FileSpec
                        Exit Function
                    End If
                On Error GoTo 0
                
                glbFSO.DeleteFile FileSpec
            
            Next Attachment
        
    End Select
    
    Response_Attachments = True

End Function

' ---------------------------------------------------------------------
'   2025-06-19 - SPOS
'
'   Stupid is acting really weird on Response to a Meeting Item (Cancel, Accept, Tenative, Decline).
'   Sometimes I get a well formed Response, sometimes a Body that is just a single No Break Space.
'   Can't figure out what the cause is. This is just a kluge, not a fix.
'
'   Discovered by accident that if I start a new Responce (and cancel his Responce) that it always works.
'   So that's what all of this is for. To handle a corner case that I almost never use.
'
' ---------------------------------------------------------------------
'
Private Function Response_MeetingItem(ByVal Original As Object, ByRef Cancel As Boolean, ByVal ResponseType As Integer) As Boolean
Const ThisProc = "Response_MeetingItem"
Response_MeetingItem = False

    '   Create a new Response
    '
    Dim oResponse As Object
    On Error Resume Next
    
        Select Case ResponseType
            Case glbResponse_Reply
                Set oResponse = Original.Reply
            Case glbResponse_ReplyAll
                Set oResponse = Original.ReplyAll
            Case glbResponse_Forward
                Set oResponse = Original.Forward
            Case Else
                Stop: GoTo Exit_Function
        End Select
        If Not Err.Number = glbError_None Then Stop: GoTo Exit_Function
        
    On Error GoTo 0
    
    '   Get the Response InspectorRec
    '
    Dim ResponseInspectorRec As Inspector_InspectorRec
    If Not Inspector_RecFromItem(oResponse, ResponseInspectorRec) Then Stop: GoTo Exit_Function
    
    With ResponseInspectorRec
    
        '   Get the Response WordDoc
        '
        Dim wDoc As Word.Document
        Set wDoc = .oInspector.WordEditor
        If wDoc Is Nothing Then Stop: GoTo Exit_Function
        
    End With
    
    '   At this point the Body is a single para with ^l line breaks.
    '   Add two ^p at the top to make Cleanup happy.
    '
    wDoc.Content.Paragraphs.Add
    wDoc.Content.Paragraphs.Add
        
    '   Show the new Response
    '   Cancel the pending Response
    '
    ResponseInspectorRec.oInspector.Activate
    Cancel = True
    
Response_MeetingItem = True
Exit Function

    '   Error Exit - Clean up any Inspector
    '
Exit_Function:

    If Not Inspector_RecInspectorCloseIfNew(ResponseInspectorRec) Then Stop: Exit Function
    
End Function
