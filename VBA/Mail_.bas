Attribute VB_Name = "Mail_"
Option Explicit
Option Private Module

'   Has Item Been Sent?
'
Public Function Mail_IsSent(ByVal Item As Object) As Boolean
Const ThisProc = "Mail_IsSent"

    Mail_IsSent = False

    If Item.ItemProperties.Item("Sent") Is Nothing Then Exit Function
    
    '   SPOS - Even with the above check, Stupid will still try and access "Sent" when it doesn't exist.
    '   So we Error trap a glbError_PropertyNotFound
    '
    On Error Resume Next
    
        If Not Item.Sent Then Exit Function
        If Err.Number = glbError_PropertyNotFound Then Exit Function
        If Err.Number <> glbError_None Then Stop: Exit Function
    
    On Error GoTo 0
    
    Mail_IsSent = True

End Function

'   Does Item have an HTMLBody?
'
Public Function Mail_HasHTMLBody(ByVal Item As Object) As Boolean
Const ThisProc = "Mail_HasHTMLBody"
Mail_HasHTMLBody = False

    If Item.ItemProperties.Item("BodyFormat") Is Nothing Then Exit Function
    If Not (Item.BodyFormat = Outlook.olFormatHTML) Then Exit Function
    
Mail_HasHTMLBody = True
End Function

'   Is the Item RTF?
'
Public Function Mail_IsRTF(ByVal Item As Object) As Boolean
Const ThisProc = "Mail_IsRTF"
Mail_IsRTF = False

    If Item.ItemProperties.Item("BodyFormat") Is Nothing Then Exit Function
    If Not (Item.BodyFormat = Outlook.olFormatRichText) Then Exit Function
    
Mail_IsRTF = True
End Function

'   Find an open Explorer for an Item (if any)
'
Public Function Mail_FindItemExplorer(ByVal Item As Object) As Outlook.Explorer

    Set Mail_FindItemExplorer = Nothing

    ' Get my Item's folder
    '
    Dim ItemFolder As Outlook.Folder
    If Not Folders_Item(Item, ItemFolder) Then Stop: Exit Function
    
    ' Get a collection of open Explorers
    '
    Dim Explorers As Outlook.Explorers
    Set Explorers = Application.Explorers
    
    ' Search Explorers for one that is open to my Item's folder
    '
    Dim ExplorerFound As Boolean: ExplorerFound = False
    Dim Explorer As Outlook.Explorer
    
    For Each Explorer In Explorers
        If Explorer.CurrentFolder.FolderPath = ItemFolder.FolderPath Then
            ExplorerFound = True
            Exit For
        End If
    Next
    
    If ExplorerFound Then Set Mail_FindItemExplorer = Explorer

End Function

'   Count the number of open Explorers for an Item (if any)
'
Public Function Mail_CountItemExplorers(ByVal Item As Object) As Long

    Mail_CountItemExplorers = 0

    ' Get my Item's folder
    '
    Dim ItemFolder As Outlook.Folder
    If Not Folders_Item(Item, ItemFolder) Then Stop: Exit Function
    
    ' Get a collection of open Explorers
    '
    Dim Explorers As Outlook.Explorers
    Set Explorers = Application.Explorers
    
    ' Search Explorers for all that are open to my Item's folder
    '
    Dim Explorer As Outlook.Explorer
    For Each Explorer In Explorers
        If Explorer.CurrentFolder.FolderPath = ItemFolder.FolderPath Then
            Mail_CountItemExplorers = Mail_CountItemExplorers + 1
        End If
    Next

End Function

'   Is a Reply/Forward?
'
'   Can not use PR_LAST_VERB_EXECUTED during a Reply/Fwd Event
'   because that won't be set on the Original until AFTER the response is sent.
'
'   All Response will have InternetReplyID because I force it in
'   Response_ForceReply. (Stupid doesn't always add this Property or set it to a value).
'
Public Function Mail_IsResponse(ByVal Item As Object) As Boolean
Mail_IsResponse = False

    '   Must have a InternetReplyID
    '
    Dim InternetReplyID As String
    If Not Misc_OLGetProperty(Item, glbPropTag_InternetReplyID, InternetReplyID) Then InternetReplyID = ""
    If InternetReplyID = "" Then Exit Function
    
    '   And must be unsent
    '
    If Mail_IsSent(Item) Then Exit Function

Mail_IsResponse = True
End Function

'   Is a Meeting Responce (Cancel, Accept, Tenative, Decline)?
'
Public Function Mail_IsMeetingResponse(ByVal Item As Object) As Boolean
Mail_IsMeetingResponse = True

    Select Case Item.Class
        Case Outlook.olMeetingCancellation
            Exit Function
        Case Outlook.olMeetingForwardNotification
            Exit Function
        Case Outlook.olMeetingResponseNegative
            Exit Function
        Case Outlook.olMeetingResponsePositive
            Exit Function
        Case Outlook.olMeetingResponseTentative
            Exit Function
    End Select

Mail_IsMeetingResponse = False
End Function

'   Is a Meeting Request (Invite, Cancel, Forward)?
'
Public Function Mail_IsMeetingRequest(ByVal Item As Object) As Boolean
Mail_IsMeetingRequest = True

    Select Case Item.Class
        Case Outlook.olMeetingCancellation
            Exit Function
        Case Outlook.olMeetingForwardNotification
            Exit Function
        Case Outlook.olMeetingRequest
            Exit Function
    End Select

Mail_IsMeetingRequest = False
End Function

'   Is a Meeting Response Action (Declined, Accepted, Tenative)?
'
Public Function Mail_IsMeetingResponseAction(ByVal Item As Object) As Boolean
Mail_IsMeetingResponseAction = True

    Select Case Item.Class
        Case Outlook.olMeetingResponseNegative
            Exit Function
        Case Outlook.olMeetingResponsePositive
            Exit Function
        Case Outlook.olMeetingResponseTentative
            Exit Function
    End Select

Mail_IsMeetingResponseAction = False
End Function


'   Get the SaveSentMessageFolder (if any)
'
Public Function Mail_GetSaveSentFolder(ByVal Item As Object, ByRef Folder As Outlook.Folder) As Boolean
Const ThisProc = "Mail_GetSaveSentFolder"

    Mail_GetSaveSentFolder = True
    Set Folder = Nothing
    
    If Item.ItemProperties.Item("SaveSentMessageFolder") Is Nothing Then Exit Function
    
    On Error Resume Next
    
        If Item.SaveSentMessageFolder Is Nothing Then Exit Function
        If Err.Number = glbError_PropertyNotFound Then Exit Function
        If Err.Number <> glbError_None Then Stop: Exit Function
        
    On Error GoTo 0
        
    Set Folder = Item.SaveSentMessageFolder

End Function

'   Get all the Items in a Folder with a specific DownloadState
'
' ---------------------------------------------------------------------
'
'   SPOS - The Item.DownloadState is not a queryable field so no Filter query possible on that.
'
'   SPOS - The values in IMAP Staus (HText Status) field are NOT the same as OlDownloadState
'   (Header Status) enumeration for an Item. It APPEARS that for IMAP Staus olFullItem = 0
'   and olHeaderOnly = 1. You have to flip them in a query.
'
'   SPOS - The IMAP Staus (HText Status) field is unreliable. It undercounts.
'
'   So we fall back on walking the Folder item by item and checking the Item.OlDownloadState. But (wait for it) ...
'
'   SPOS - The Item.DownloadState doesn't always update until some time after the Header download.
'
'   Add to all of this the whole "foce add the HS (Header Status) column to a view on delete" and you have to wonder -
'
'                   How many different ways can they fuck up one field?
'
' ---------------------------------------------------------------------
'
Public Function Mail_SearchDownloadState(ByVal DownloadState As Integer, ByVal Folder As Outlook.Folder, ByVal Results As Collection) As Boolean
Mail_SearchDownloadState = False
    
    Dim Item As Object
    For Each Item In Folder.Items
        If Item.DownloadState = DownloadState Then Results.Add Item
    Next Item

    If Results.count = 0 Then Exit Function

Mail_SearchDownloadState = True
End Function

Public Function Mail_NewMail() As Boolean
Mail_NewMail = False

    Dim NewMail As Outlook.MailItem
    Set NewMail = Application.CreateItem(Outlook.olMailItem)
    NewMail.GetInspector.Activate
    
Mail_NewMail = True
End Function

'   Show the Address Book
'
Public Function Mail_AddressBook() As Boolean
Mail_AddressBook = False
    
    '   If the Active Window supports the Address Book command - do it and done
    '
    If Ribbon_ExecuteMSO(ActiveWindow, glbidMSO_AddressBook) Then Mail_AddressBook = True:  Exit Function
    
    '   Else - Jump to a Home Inbox Explorer (if any) and do it from there
    '
    If Utility_HomeInboxExplorer() Is Nothing Then Exit Function
    If Not Ribbon_ExecuteMSO(ActiveWindow, glbidMSO_AddressBook) Then Exit Function
    
Mail_AddressBook = True
End Function

'   Open a new mail and the address book
'
Public Function Mail_NewMailAndAddressBook() As Boolean
Mail_NewMailAndAddressBook = False

    If Not Mail_NewMail() Then Stop: Exit Function
    If Not Mail_AddressBook() Then Stop: Exit Function

Mail_NewMailAndAddressBook = True
End Function

'   Clear the To, CC, and BCC fields on the Current Inspector Item
'
'       To clear other address fields you have to use OutlookSpy IMAPI mode.
'
Public Sub Mail_ClearAddresses()

    Dim Item As Outlook.MailItem
    If Not Misc_GetActiveItem(Item, InspectorOnly:=True) Then Exit Sub
    
    Item.CC = ""
    Item.BCC = ""
    Item.To = ""

End Sub

' ---------------------------------------------------------------------
'   get Address Strings. e.g. "Joe Doe <john@company.com> ; ... "
' ---------------------------------------------------------------------

'   Get an Address String from a Name and eMail
'
Public Function Mail_AdrString(ByVal NameAdr As String, ByVal eMailAdr As String) As String

    Dim Name As String
    Name = NameAdr
    
    Dim eMail As String
    eMail = eMailAdr
    
    If Name = eMail Then Name = ""
    If eMail <> "" Then eMail = "<" & eMail & ">"
    Mail_AdrString = Trim(Name & " " & eMail)

End Function

'   Get an AdrString of Recipients from a Recipients Collection (Optional Filter by RecipientType)
'
Public Function Mail_AdrStringRecipients(ByVal oRecipients As Outlook.Recipients, Optional ByVal RecipientType As Long = -1) As String

    Dim sRecipients As String
    Dim oRecipient As Outlook.Recipient
    Dim sRecipient As String
    
    For Each oRecipient In oRecipients
        
        If (RecipientType = -1) Or (oRecipient.Type = RecipientType) Then
            sRecipient = Mail_AdrString(oRecipient.Name, oRecipient.Address)
            If sRecipient <> "" Then
                If sRecipients <> "" Then sRecipients = sRecipients & " ; "
                sRecipients = sRecipients & sRecipient
            End If
        End If
    
    Next oRecipient
    Mail_AdrStringRecipients = sRecipients
    
End Function

' ---------------------------------------------------------------------
'   Contacts Lookup
' ---------------------------------------------------------------------

'   Get DisplayAs from Contacts for an eMail address
'
Public Function Mail_eMailDisplayAs(ByVal eMailAdr As String, ByRef DisplayAs As String) As Boolean
Mail_eMailDisplayAs = False

    Dim oFolder As Outlook.Folder
    Set oFolder = Folders_KnownPath(glbKnownPath_Contacts)
    
    Dim Filter As String
    Filter = _
                 "([Email1Address] = '" & eMailAdr & "')" & _
        " or " & "([Email2Address] = '" & eMailAdr & "')" & _
        " or " & "([Email3Address] = '" & eMailAdr & "')"

    Dim oContact As Outlook.ContactItem
    Set oContact = oFolder.Items.Find(Filter)
    If oContact Is Nothing Then Exit Function
    
    With oContact
    Select Case eMailAdr
        Case .Email1Address
            DisplayAs = .Email1DisplayName
        Case .Email2Address
            DisplayAs = .Email2DisplayName
        Case .Email3Address
            DisplayAs = .Emai31DisplayName
        Case Else
            Stop
    End Select
    End With

Mail_eMailDisplayAs = True
End Function

' =====================================================================
'   Item.BillingInformation
' =====================================================================
'
Private Function Mail_BIArray(ByVal Item As Object, ByRef BIData As Variant) As Boolean
Mail_BIArray = False

    '   Make sure the Item has a BillingInformation field and get it
    '
    Dim BI As String
    On Error Resume Next
        BI = Item.BillingInformation
        If Err.Number <> glbError_None Then Stop: Exit Function
    On Error GoTo 0
    
    '   If BI doesn't start with my BI SIG - Init
    '
    If Left(BI, glbGUIDLen) <> glbBI_Sig Then BI = glbBI_Sig & String(glbBI_UBound, glbBI_Sep)
    
    '   Split the BI and return the array
    '
    BIData = Split(BI, glbBI_Sep)
    If UBound(BIData) <> glbBI_UBound Then Stop: Exit Function

Mail_BIArray = True
End Function

Public Function Mail_BIGet(ByVal Item As Object, ByVal BIInx As Long, ByRef value As String) As Boolean
Mail_BIGet = False

    If BIInx > glbBI_UBound Then Stop: Exit Function
    
    Dim BIData As Variant
    If Not Mail_BIArray(Item, BIData) Then Stop: Exit Function
    value = BIData(BIInx)

Mail_BIGet = True
End Function

Public Function Mail_BISet(ByVal Item As Object, ByVal BIInx As Long, ByVal value As String) As Boolean
Mail_BISet = False

    If BIInx > glbBI_UBound Then Stop: Exit Function

    Dim BIData As Variant
    If Not Mail_BIArray(Item, BIData) Then Stop: Exit Function
    BIData(BIInx) = value
    Item.BillingInformation = Join(BIData, glbBI_Sep)

Mail_BISet = True
End Function

' =====================================================================
'   Export an Object as a PDF file
'
'   !!  All exits MUST be thru Goto Exit_Sub !!
'
' =====================================================================
'
Public Sub Mail_ExportAsPDF()
Const ThisProc = "Mail_ExportAsPDF"

    Dim Export As Object

    '   Get a Single Selected Item from the current Inspector or Explorer
    '
    Dim Original As Object
    If Not Misc_GetActiveItem(Original, ThisProc) Then GoTo Exit_Sub
    
    '   Get an InspectorRec for the Original
    '
    Dim OriginalInspectorRec As Inspector_InspectorRec
    If Not Inspector_RecFromItem(Original, OriginalInspectorRec) Then Stop: GoTo Exit_Sub
    
    With OriginalInspectorRec
    
        '   Get the Original wDoc - error if none
        '
        Dim wDocOriginal As Word.Document
        Set wDocOriginal = .oInspector.WordEditor
        If wDocOriginal Is Nothing Then
            Msg_Box Proc:=ThisProc, Step:="Has WordEditor?", Text:="The selected Item does not support a WordEditor."
            GoTo Exit_Sub
        End If
    
    End With
    
    '   Clone the Original as Export in DeletedItems
    '   (Next Export.Delete will delete it permanently)
    '
    If Not File_CloneInDeleted(Original, Export) Then Stop: GoTo Exit_Sub
    
    '   If needed - convert Plain Text to HTML
    '
    If Not Format_Plain2HTML(Export) Then GoTo Exit_Sub
    
    '   Get an InspectorRec for the Export
    '
    Dim ExportInspectorRec As Inspector_InspectorRec
    If Not Inspector_RecFromItem(Export, ExportInspectorRec) Then Stop: GoTo Exit_Sub
    
    With ExportInspectorRec
    
        '   Get the Export wDoc
        '
        Dim wDocExport As Word.Document
        Set wDocExport = .oInspector.WordEditor
        If wDocExport Is Nothing Then Stop: GoTo Exit_Sub
    
    End With
    
    '   Unprotect Export
    '
    If wDocExport.ProtectionType <> wdNoProtection Then wDocExport.UnProtect
        
    '   Insert three blank para
    '   (The first two are to keep HotRod format from mucking with the thrid)
    '
    wDocExport.Paragraphs.Add wDocExport.Range
    wDocExport.Paragraphs.Add wDocExport.Range
    wDocExport.Paragraphs.Add wDocExport.Range
    
    '   Format the Header Para(3) and add a Footer
    '
    Dim HPara As Word.Paragraph
    Set HPara = wDocExport.Paragraphs.Item(3)
    Dim HParaRange As Word.Range
    Set HParaRange = HPara.Range.Duplicate
    
    Format_PlainAll HParaRange
    HPara.Format.SpaceAfter = 8
    HPara.Format.Borders.Item(wdBorderTop).LineStyle = wdLineStyleSingle
    HPara.Format.Borders.Item(wdBorderTop).LineWidth = wdLineWidth300pt
    HParaRange.Font.Name = "Tahoma"
    
    Format_PlainFooter wDocExport.Sections.Item(1).Footers.Item(wdHeaderFooterPrimary).Range, "Retrieved "
    
    '   Setup for Gosub Header_ calls
    '
    Dim HRange As Word.Range
    Set HRange = HParaRange.Duplicate
    HRange.Collapse wdCollapseStart
    Dim Tag As String
    Dim Text As String
    
    '   Gosub Header_ based on the Item Type
    '
    Select Case True
        Case TypeOf Original Is Outlook.MailItem
            GoSub Header_Mail
        Case TypeOf Original Is Outlook.MeetingItem
            GoSub Header_Invite
        Case Else
            GoSub Header_Other
    End Select
    
    '   Remove the last ^l from the Header Para so it ends with just a ^p
    '
    Set HRange = HParaRange.Duplicate
    HRange.Start = HRange.End - 2
    HRange.Collapse wdCollapseStart
    HRange.Delete wdCharacter, 1
        
    '   Remove the two blank Para from the top
    '
    wDocExport.Paragraphs.First.Range.Delete
    wDocExport.Paragraphs.First.Range.Delete
    
    '   Build a FileSpec from the Subject and a Timestamp.
    '
    Dim FileSpec As String
    FileSpec = glbExportAsPDFPath & Export.Subject & " " & Misc_NowStamp() & ".pdf"
    FileSpec = File_CleanupFileSpec(FileSpec)
    
    '   Export As PDF
    '
    wDocExport.ExportAsFixedFormat _
        OutputFileName:=FileSpec, _
        ExportFormat:=wdExportFormatPDF, _
        OptimizeFor:=wdExportOptimizeForPrint

    '   Close and Delete Export (as soon as possible)
    '   Put the FileSpec on the Clipboard
    '
    If Not Inspector_RecInspectorCloseIfNew(ExportInspectorRec) Then Stop: GoTo Exit_Sub
    If Not File_ItemDelete(Export) Then Stop: GoTo Exit_Sub
    Misc_ClipSet FileSpec

    '   Show and Tell
    '
    Msg_Box Proc:=ThisProc, Step:="Show and Tell", Icon:=vbInformation, _
    Text:="PDF Export file created and the FileSpec copied to the Clipboard."

Exit_Sub:

    '   Cleanup
    '
    If Not Inspector_RecInspectorCloseIfNew(OriginalInspectorRec) Then Stop: Exit Sub
    If Not Inspector_RecInspectorCloseIfNew(ExportInspectorRec) Then Stop: Exit Sub
    If Not File_ItemDelete(Export) Then Stop: Exit Sub

Exit Sub
    
' ---------------------------------------------------------------------
'   Header Build Gosubs
' ---------------------------------------------------------------------

Header_Other:

    Tag = "Subject: "
    Text = Original.Subject
    GoSub Header_AddLine

Return

Header_Mail:

    Tag = "From: "
    Text = Mail_AdrString(Original.SenderName, Original.SenderEmailAddress)
    GoSub Header_AddLine
    
    Tag = "Sent: "
    Text = Format(Original.SentOn, "YYYY-MM-DD hh:mm")
    GoSub Header_AddLine
    
    Tag = "To: "
    Text = Mail_AdrStringRecipients(Original.Recipients, Outlook.olTo)
    GoSub Header_AddLine
    
    Tag = "Cc: "
    Text = Mail_AdrStringRecipients(Original.Recipients, Outlook.olCC)
    GoSub Header_AddLine
    
    Tag = "Subject: "
    Text = Original.Subject
    GoSub Header_AddLine

Return

Header_Invite:

    Tag = "From: "
    Text = Mail_AdrString(Original.SenderName, Original.SenderEmailAddress)
    GoSub Header_AddLine
    
    Tag = "Sent: "
    Text = Format(Original.SentOn, "YYYY-MM-DD hh:mm")
    GoSub Header_AddLine
    
    Tag = "Organizer: "
    Text = Mail_AdrStringRecipients(Original.Recipients, Outlook.olOrganizer)
    GoSub Header_AddLine
    
    Tag = "Required: "
    Text = Mail_AdrStringRecipients(Original.Recipients, Outlook.olRequired)
    GoSub Header_AddLine
    
    Tag = "Optional: "
    Text = Mail_AdrStringRecipients(Original.Recipients, Outlook.olOptional)
    GoSub Header_AddLine
    
    Tag = "Subject: "
    Text = Original.Subject
    GoSub Header_AddLine

'    Tag = "Location: "
'    Text = Original.Location
'    GoSub Header_AddLine

'    Tag = "When: "
'    Text = Original.Subject
'    GoSub Header_AddLine

Return

Header_AddLine:

    HRange.InsertAfter Tag
    HRange.Font.Bold = True
    HRange.Collapse wdCollapseEnd
    HRange.InsertAfter Text
    HRange.Font.Bold = False
    HRange.Collapse wdCollapseEnd
    HRange.InsertAfter Chr(11)
    HRange.Collapse wdCollapseEnd

Return

End Sub

