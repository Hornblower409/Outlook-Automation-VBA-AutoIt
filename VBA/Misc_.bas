Attribute VB_Name = "Misc_"
Option Explicit
Option Private Module

' =====================================================================
'   Zombie
' =====================================================================

'   Is an Object Reference to an Item that has been moved or deleted?
'
Public Function Misc_ItemIsZombie(ByVal Item As Object) As Boolean
Misc_ItemIsZombie = True

    On Error Resume Next
    
        If Item Is Nothing Then Exit Function
        
        '   SPOS - No way to check for "<The item has been moved or deleted.>"
        '   or <Automation Error> which is what each field in the Item is really
        '   if it was deleted. So instead we access a a field we know every item has
        '   (EntryId) and if throws an error, we know it's gone.
        '
        Dim Dummy As String
        Dummy = Item.EntryId
        If Err.Number <> glbError_None Then Exit Function
        
    On Error GoTo 0
    
    '   2025-01-15 - Commented Out.
    '
    '   - Was stopping me from moving (archiving) Appointment/Meeting to Projects.
    '   - I think this was a leftover from Is It Gone.
    '   - Or I changed the Default Misc_ItemIsZombie and didn't change this.
    '
    '   If it's still alive but an AppointmentItem - good enough.
    '
    '       If the Item is an Appointment/Meeting occurances of a Recuring
    '       then the Item I have is the Parent Series. The Delete only removed
    '       the occurance, the Item I'm looking at (the Series) is still alive.
    '
    ' If (TypeOf Item Is Outlook.AppointmentItem) Then Exit Function

Misc_ItemIsZombie = False
End Function

' =====================================================================
'   Storage Size String
' =====================================================================

'   Convert bytes into storage size string
'
'       https://stackoverflow.com/questions/39912582/converting-file-sizes-from-bytes-to-kb-or-mb
'
Public Function Misc_BytesToStr(ByVal SizeInBytes As Double) As String

    Dim Size_Bytes As Double
    Size_Bytes = SizeInBytes
    
    Dim Ts() As String
    ReDim Ts(4)
    Ts(0) = "bytes"
    Ts(1) = "KB"
    Ts(2) = "MB"
    Ts(3) = "GB"
    Ts(4) = "TB"

    Dim Size_Counter As Integer
    Size_Counter = 0

    If Size_Bytes <= 1 Then
        Size_Counter = 1
    Else
        While Size_Bytes > 1
            Size_Bytes = Size_Bytes / 1024
            Size_Counter = Size_Counter + 1
        Wend
    End If

    Misc_BytesToStr = Format(Size_Bytes * 1024, "##0.0#") & " " & Ts(Size_Counter - 1)
    
End Function

' =====================================================================
'   Hex Encode/Decode
' =====================================================================

'   Would appear that Outlook stores chars in a string as two byte little endian.
'
Public Function Misc_PlainToHex(ByVal Plain As String, ByRef HexStr As String) As Boolean
Const ThisProc = "Misc_PlainToHex"

    Misc_PlainToHex = False
    
    HexStr = ""
    If LenB(Plain) = 0 Then Exit Function
    
    Dim iX As Long
    For iX = 1 To LenB(Plain)
      HexStr = HexStr & Right("00" & Hex(AscB(MidB(Plain, iX, 1))), 2)
    Next iX
    
    Misc_PlainToHex = True

End Function

Public Function Misc_HexToPlain(ByRef Plain As String, ByVal HexStr As String) As Boolean
Const ThisProc = "Misc_HexToPlain"
Misc_HexToPlain = False
    
    If LenB(HexStr) = 0 Then
        Msg_Box Proc:=ThisProc, Step:="Check Args", Text:="HexStr can not be an empty string."
        Exit Function
    End If
    
    If (LenB(HexStr) Mod 2) <> 0 Then
        Msg_Box Proc:=ThisProc, Step:="Check Args", Text:="HexStr lenght must be multiples of two."
        Exit Function
    End If
    
    Plain = ""
    Dim iX As Long
    For iX = 1 To Len(HexStr) Step 2
    
        Plain = Plain & ChrB(CLng("&H" & Mid(HexStr, iX, 2)))
    
    Next iX
    
Misc_HexToPlain = True
End Function

' =====================================================================
'   Environment Variables
' =====================================================================

'   Get the value of an Environment variable
'
'       Do NOT put "%" around the name.
'
Public Function Misc_EnvironmentGet(ByVal EnvName As String, ByRef EnvValue As String) As Boolean
Const ThisProc = "Misc_EnvironmentGet"

    Misc_EnvironmentGet = False
    
    EnvValue = Environ(EnvName)
    If EnvValue = "" Then
        Msg_Box Proc:=ThisProc, Text:="Environment Variable " & EnvName & " does not exist."
        Exit Function
    End If
    
    Misc_EnvironmentGet = True

End Function

' =====================================================================
'   User Defined Error
' =====================================================================

'   Raise a User Defined Error
'
'       ErrorDef = "NNNN {Description}" e.g. "0515 cInspector Collection is Empty" (see Globals_ Private Errors)
'
'       From: https://bettersolutions.com/vba/error-handling/raising-errors.htm
'
Public Sub Misc_ErrorRaise( _
    ByVal ErrorDef As String, _
    Optional ByVal Proc As String = "", _
    Optional ByVal Step As String = "", _
    Optional ByVal Param As String = "")
    
    '   Generate a Private Error Number
    '
    Dim ErrorNum As Long
    ErrorNum = CLng(Left(ErrorDef, 4))
    If ErrorNum < 500 Then Stop: Exit Sub
    ErrorNum = -10000000 + ErrorNum
    
    '   Build the Error Source String
    '
    Dim ErrSource As String
    ErrSource = Proc & ", " & Step
    
    If Proc <> "" And Step <> "" Then ErrSource = "Proc: " & Proc & ", Step: " & Step
    If Proc <> "" And Step = "" Then ErrSource = "Proc: " & Proc
    If Proc = "" And Step <> "" Then ErrSource = "Step: " & Step
    If ErrSource <> "" Then ErrSource = "Source = " & ErrSource
    
    '   Build the Error Param String
    '
    Dim ErrParam As String: ErrParam = ""
    If Param <> "" Then ErrParam = "Param = '" & Param & "'"
    
    '   Get the Error Text from the Error Definition
    '
    Dim ErrText As String: ErrText = ""
    ErrText = Mid(ErrorDef, 6)
    If ErrText = "" Then Stop: Exit Sub
    
    '   Build the full Error Description
    '
    '       SPOS - Stupid doesn't show the Source param in the Error Dialog.
    '       So I add it to the description.
    '
    Dim ErrDesc As String: ErrDesc = ""
    If ErrSource <> "" Then ErrDesc = ErrDesc & ErrSource & glbBlankLine
    If ErrParam <> "" Then ErrDesc = ErrDesc & ErrParam & glbBlankLine
    ErrDesc = ErrDesc & ErrText
    
    '   Raise the Roof
    '
    Err.Raise Number:=ErrorNum, Source:=ErrSource, Description:=ErrDesc
    
End Sub

' =====================================================================
'   AltVBA
' =====================================================================

'   Running from an AltVBA?
'
Public Function Misc_UsingAltVBA(Optional ByVal ShowStartupMsg As Boolean = False) As Boolean
Const ThisProc = "Misc_UsingAltVBA"
Misc_UsingAltVBA = False

    Dim AltVBAEnvValue As String
    AltVBAEnvValue = Environ(glbAltVBAEnv)
    If AltVBAEnvValue = "" Then Exit Function
    
    If ShowStartupMsg Then
        Msg_Box Proc:=ThisProc, Step:="Check AltVBA", _
                Text:="Running AltVBA = '" & AltVBAEnvValue & "'." & glbBlankLine & _
                      "Application Startup terminated."
    End If

Misc_UsingAltVBA = True
End Function

' =====================================================================
'   PropertyAccessor
' =====================================================================

' ---------------------------------------------------------------------
'   Outlook MAPI Property Accessors
'
'       SPOS - Hoarked when the Get Value is a field on a form.
'
'           e.g. Misc_OLGetProperty(Item, glbPropTag_FlagRequest, FollowUpForm.Title)
'           Will leave .Title = "" even when it has a value.
'
'           Set seems to be OK with it. Below works fine.
'           e.g. Misc_OLSetProperty(Item, glbPropTag_FlagRequest, FollowUpForm.Title)
'
'       Return FALSE if the operation fails (Property does not exist/can not be set)
'
' ---------------------------------------------------------------------
'
Public Function Misc_OLGetProperty(ByVal Item As Object, ByVal PropTag As String, ByRef value As Variant) As Boolean
Misc_OLGetProperty = False

    Dim PA As Outlook.PropertyAccessor
    
    On Error GoTo ErrExit
    
        Set PA = Item.PropertyAccessor
        value = PA.GetProperty(PropTag)

    On Error GoTo 0
    
Misc_OLGetProperty = True
ErrExit: End Function

Public Function Misc_OLSetProperty(ByVal Item As Object, ByVal PropTag As String, ByVal value As Variant) As Boolean
Misc_OLSetProperty = False

    Dim PA As Outlook.PropertyAccessor
    Set PA = Item.PropertyAccessor
    
    On Error Resume Next
    
        PA.SetProperty PropTag, value
        Select Case Err.Number
            Case glbError_None
            Case glbError_TypeMismatch
                Stop: Exit Function
            Case Else
                Exit Function
        End Select

    On Error GoTo 0
    
Misc_OLSetProperty = True
End Function

' =====================================================================
'   URL Encode/Decode
' =====================================================================

'   Decode/Encode a String with URL Percent Encoding for "\/%"
'
Public Function Misc_URLDecode(ByVal Encoded As String) As String

    '   !! You have to do % (%25) last or you could get hoarked !!
    '
    Misc_URLDecode = Encoded
    Misc_URLDecode = Replace(Misc_URLDecode, "%5C", "\", Compare:=vbTextCompare)
    Misc_URLDecode = Replace(Misc_URLDecode, "%2F", "/", Compare:=vbTextCompare)
    Misc_URLDecode = Replace(Misc_URLDecode, "%25", "%", Compare:=vbTextCompare)

End Function
Public Function Misc_URLEncode(ByVal Plain As String) As String

    '   !! You have to do % (%25) first or you could get hoarked !!
    '
    Misc_URLEncode = Plain
    Misc_URLEncode = Replace(Misc_URLEncode, "%", "%25", Compare:=vbTextCompare)
    Misc_URLEncode = Replace(Misc_URLEncode, "/", "%2F", Compare:=vbTextCompare)
    Misc_URLEncode = Replace(Misc_URLEncode, "\", "%5C", Compare:=vbTextCompare)

End Function

' =====================================================================
'   GUID
' =====================================================================

'   Random 8-4-4-4-12 GUID Generator
'
'   See https://www.linkedin.com/pulse/excel-vba-create-guids-easily-daniel-ferry/
'
'   Any changes here must change Public Const glbGUIDLen as Long = 36
'
Public Function Misc_MakeGUID() As String

    Dim Pattern As String
    Pattern = "00000000-0000-1000-A000-000000000000"
    Dim GUID As String
    GUID = ""
    
    Randomize
    Dim Index As Integer
    For Index = 1 To Len(Pattern)
    
        If Mid(Pattern, Index, 1) = "0" Then
            GUID = GUID & Hex(Rnd * 15)
        Else
            GUID = GUID & Mid(Pattern, Index, 1)
        End If
    
    Next Index
    
    Misc_MakeGUID = UCase(GUID)
    If Len(Misc_MakeGUID) <> glbGUIDLen Then Stop: Exit Function
    
End Function

' =====================================================================
'   Clipboard
' =====================================================================

'   Format
'
'       https://learn.microsoft.com/en-us/dotnet/api/system.windows.forms.textdataformat?view=windowsdesktop-8.0https://learn.microsoft.com/en-us/dotnet/api/system.windows.forms.textdataformat?view=windowsdesktop-8.0
'
'       Text    0   ANSI text
'       Unicode 1   Windows Unicode text
'       RTF     2   Rich Text Format (RTF)
'       Html    3   HTML data
'       CSV     4   Comma-Separated Calue (CSV) format
'
'   https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dataobject-object
'
'       The DataObject works like the Clipboard. If you copy a text string to a DataObject,
'       the DataObject stores the text string. If you copy a second string of the same format
'       to the DataObject, the DataObject discards the first text string and stores a copy of
'       the second string. It stores one piece of text of a specified format from the most
'       recent operation.
'
'   To Clear the Clipboard
'
'       So you don't get a "You placed a large amount of data on the Clipboard ... keep it?"
'       dialog when you exit.
'
'       https://stackoverflow.com/questions/32736915/how-to-clear-office-clipboard-with-vba
'
'           oData.SetText Text:=Empty           'Clear
'           oData.PutInClipboard

'   Put a String on the Clipboard in a specific Format
'
Public Sub Misc_ClipSet(ByVal Text As String, Optional ByVal Format As Long = 1)

    Dim oData As MSForms.DataObject
    Set oData = New MSForms.DataObject

    oData.SetText Text, Format
    oData.PutInClipboard
    Misc_ClipSetWait
    
End Sub

'   Wait after ClipSet
'
'       Give any Clipboard Handlers (e.g. Clipmate) a chance to get the clip.
'       If we don't wait, and there is another Clipboard put right after this one,
'       Clipboard Handler doesn't have time to get and stash it.
'
Public Sub Misc_ClipSetWait(Optional ByVal SleepTime As Long = 250)

    Sleep SleepTime
    DoEvents

End Sub

'   Get a String from the Clipboard in a specific Format
'
'       Returns True if there is text on the Clipboard in that Format, Else False.
'
Public Function Misc_ClipGet(ByRef value As String, Optional ByVal Format As Long = 1) As Boolean
Misc_ClipGet = False

    Dim oData As MSForms.DataObject
    Set oData = New MSForms.DataObject
    
    '   Can throw an error on an Empty Clipboard
    '
    '       Get the current Clipboard entry
    '       Clipboard entry has text in the requested Format?
    '
    On Error Resume Next
        oData.GetFromClipboard
        If Err.Number <> glbError_None Then Exit Function
        If Not oData.GetFormat(Format) Then Exit Function
        If Err.Number <> glbError_None Then Exit Function
    On Error GoTo 0
    
    '   Get the text and done
    '
    value = oData.GetText(Format)

Misc_ClipGet = True
End Function

' =====================================================================
'   Item Type
' =====================================================================

'   Get a Friendly TypeName(oItem). "Unknown" if not found.
'
'   I'm not really sure what this is or where Stupid gets it from.
'   Some get compressed e.g. Item.Class = olMeetingResponseTentative
'   returns "MeetingItem". Others are expanded e.g. "TaskRequestAcceptItem", etc.
'
'   https://learn.microsoft.com/en-us/office/vba/outlook/how-to/items-folders-and-stores/outlook-item-objects
'   https://learn.microsoft.com/en-us/office/vba/api/outlook.olobjectclass
'
Public Function Misc_ItemTypeName(ByVal oItem As Object) As String

    Dim Name As String
    Select Case TypeName(oItem)
        Case "AppointmentItem"
            Select Case oItem.MeetingStatus
                Case Outlook.olMeeting:                     Name = "Meeting"
                Case Outlook.olMeetingReceived:             Name = "Meeting (Tentative)"
                Case Outlook.olMeetingCanceled:             Name = "Meeting (Cancelled)"
                Case Outlook.olMeetingReceivedAndCanceled:  Name = "Meeting (Cancelled)"
                Case Else:                                  Name = "Appointment"
            End Select
        Case "ContactItem":                             Name = "Contact"
        Case "DistListItem":                            Name = "Distribution List"
        Case "DocumentItem":                            Name = "Document"
        Case "JournalItem":                             Name = "Journal Entry"
        Case "MailItem":                                Name = "Mail"
        Case "MeetingItem"
            Select Case oItem.Class
                Case Outlook.olMeetingCancellation:         Name = "Meeting Cancellation"
                Case Outlook.olMeetingForwardNotification:  Name = "Meeting Forwarding"
                Case Outlook.olMeetingRequest:              Name = "Meeting Invite"
                Case Outlook.olMeetingResponseNegative:     Name = "Meeting Refusal"
                Case Outlook.olMeetingResponsePositive:     Name = "Meeting Acceptance"
                Case Outlook.olMeetingResponseTentative:    Name = "Meeting Tentative Acceptance"
                Case Else:                                  Name = "Meeting Invite/Response"
            End Select
        Case "NoteItem":                                Name = "Sticky Note"
        Case "PostItem":                                Name = "Post"
        Case "RemoteItem":                              Name = "Remote Mail Header"
        Case "ReportItem":                              Name = "Mail Delivery Report"
        Case "SharingItem":                             Name = "Sharing Invite"
        Case "StorageItem":                             Name = "Hidden Storage"
        Case "TaskItem":                                Name = "Task"
        Case "TaskRequestAcceptItem":                   Name = "Task Accept"
        Case "TaskRequestDeclineItem":                  Name = "Task Decline"
        Case "TaskRequestItem":                         Name = "Task Assignment"
        Case "TaskRequestUpdateItem":                   Name = "Task Update"
        Case Else:                                      Name = "Unknown"
    End Select
    
    Misc_ItemTypeName = Name

End Function

' =====================================================================
'   Active Item
' =====================================================================

'   Get the Active Item
'
Public Function Misc_GetActiveItem( _
    ByRef oItem As Object, _
    Optional ByVal CallingProc As String = "", _
    Optional ByVal InspectorOnly As Boolean = False, _
    Optional ByVal ExplorerOnly As Boolean = False, _
    Optional ByVal SavedOnce As Boolean = False, _
    Optional ByVal Saved As Boolean = False _
    ) As Boolean
    
Const ThisProc = "Misc_GetActiveItem"
Misc_GetActiveItem = False

    '   Who shall I say is calling?
    '
    Dim ProcName As String
    ProcName = ThisProc
    If CallingProc <> "" Then ProcName = CallingProc
    
    '   Get an Item from where ever
    '
    Select Case True
    
        Case InspectorOnly
        
            If Not (TypeOf ActiveWindow Is Outlook.Inspector) Then
                Msg_Box Proc:=ProcName, _
                Text:="The Active Window must be an Inspector."
                Exit Function
            End If
            Set oItem = ActiveInspector.CurrentItem
        
        Case ExplorerOnly
        
            If Not (TypeOf ActiveWindow Is Outlook.Explorer) Then
                Msg_Box Proc:=ProcName, _
                Text:="The Active Window must be an Explorer."
                Exit Function
            End If
            If Not (ActiveExplorer.Selection.count = 1) Then
                Msg_Box Proc:=ProcName, _
                Text:="The Active Window must be an Explorer with only a single Item selected."
                Exit Function
            End If
            Set oItem = ActiveExplorer.Selection.Item(1)
        
        Case Else ' Either
            
            Select Case True
            
                Case (TypeOf ActiveWindow Is Outlook.Inspector)
                
                    Set oItem = ActiveInspector.CurrentItem
                
                Case (TypeOf ActiveWindow Is Outlook.Explorer)
                
                    If Not (ActiveExplorer.Selection.count = 1) Then
                        Msg_Box Proc:=ProcName, _
                        Text:="The Active Explorer must have only a single Item selected."
                        Exit Function
                    End If
                    Set oItem = ActiveExplorer.Selection.Item(1)

                Case Else ' Can this happen?
                
                    Stop: Exit Function
                    
            End Select
        
    End Select
    
    '   Check for states of Grace
    '
    Select Case True
    
        Case SavedOnce
        
            If Len(oItem.EntryId) = 0 Then
                Set oItem = Nothing
                Msg_Box Proc:=ProcName, _
                Text:="The Active Item must be have been saved at least once."
                Exit Function
            End If
            
        Case Saved
        
            If Not oItem.Saved Then
                Set oItem = Nothing
                Msg_Box Proc:=ProcName, _
                Text:="The Active Item must be saved."
                Exit Function
            End If
        
        Case Else ' Continue
            
    End Select

Misc_GetActiveItem = True
End Function

' =====================================================================
'   Item Inspect
' =====================================================================

Public Sub Misc_ItemInspect()

    Dim oItem As Object
    Misc_GetActiveItem oItem
    
    Stop
    
End Sub

' =====================================================================
'   Now
' =====================================================================

Public Function Misc_NowString() As String

    Dim dNow As Date: dNow = CDate(Now)
    Misc_NowString = Format(dNow, "yyyy-mm-dd") & " " & Format(dNow, "hh:mm:ss")

End Function

Public Function Misc_NowStamp() As String

    Dim dNow As Date
    Dim msNow As String
    
    msNow = Timer: dNow = CDate(Now)
    Misc_NowStamp = Format(dNow, "yyyy-mm-dd") & "_" & Format(dNow, "hhmmss") & Right(Format(msNow, "0.00"), 2)
    
End Function

' =====================================================================
'   EntryId and StoreId
' =====================================================================

'   Get an Item from it's Entry and (optional) Store IDs
'
Public Function Misc_GetItemFromID(ByVal EntryId As String, Optional ByVal StoreID As String = "") As Object
Set Misc_GetItemFromID = Nothing

    On Error Resume Next
    
        If StoreID <> "" Then
            Set Misc_GetItemFromID = Session.GetItemFromID(EntryId, StoreID)
        Else
            Set Misc_GetItemFromID = Session.GetItemFromID(EntryId)
        End If
        
        Select Case Err.Number
            Case glbError_None
            Case glbError_ObjectNotFound
            Case glbError_MsgInterfaceUnknown
            Case Else
                Stop: Exit Function
        End Select
    
    On Error GoTo 0

End Function

' =====================================================================
'   Arrays
' =====================================================================

'   2025-05-30 - NOT FULLY TESTED
'
'   Find a Value in a 1D Array
'
Public Function Misc_ArrayFind(ByRef aArray As Variant, ByVal value As Variant) As Long
Misc_ArrayFind = -1

    If Misc_ArrayIsNothing(aArray) Then Exit Function
    
    Dim RowIx As Long
    For RowIx = LBound(aArray) To UBound(aArray)
        If aArray(RowIx) = value Then
            Misc_ArrayFind = RowIx
            Exit Function
        End If
    Next RowIx

End Function

'   Insert a new element into a 1D Variant Array
'
'   Index = 0  -> Prefix to the top
'   Index = -1 -> Append to the end
'   Index = 1  -> Insert after the first element
'
'   From https://stackoverflow.com/questions/42339959/how-to-insert-a-new-value-at-certain-index-in-an-array-and-shift-everything-down
'
Public Function Misc_ArrayInsert(ByRef aArray As Variant, ByVal vElement As Variant, Optional ByVal IndexLocation As Long = -1) As Boolean
Misc_ArrayInsert = False

    Dim Index As Long
    Index = IndexLocation
    
    '   Check the Index is in range
    '
    If (Index < -1) Or (Index > UBound(aArray) + 1) Then Stop: Exit Function
    
    '   ReDim Preserve to one larger than it is now.
    '   If Index = -1 make it the new last element
    '
    ReDim Preserve aArray(0 To UBound(aArray) + 1)
    If Index = -1 Then Index = UBound(aArray)
    
    '   Walk the Array backwards and shift everything down one
    '   till we get to the insertion point
    '
    Dim Inx As Long
    For Inx = UBound(aArray) To Index + 1 Step -1
        aArray(Inx) = aArray(Inx - 1)
    Next Inx
    
    '   Put the new element in the gap
    '
        aArray(Index) = vElement
    
Misc_ArrayInsert = True
End Function

'   Returns a copy of a dynamic string array with any empty ("") elements removed.
'
'       If there are no non-empty elemets in the original array it returns an unallocated array.
'       Join has no problem with this and returns an empty string (no delims).
'
Public Function Misc_ArrayCompress(ByRef aArray() As String) As String()
   
    '   If aArray Is Unallocated/Inaccessible - Bail
    '
    If Misc_ArrayIsNothing(aArray) Then Exit Function
    
    '   bArray = empty array the same size as aArray
    '
    Dim bArray() As String
    ReDim bArray(LBound(aArray) To UBound(aArray))

    '   Copy A to B skipping any empties
    '
    Dim aInx As Long
    Dim bInx As Long: bInx = LBound(bArray)
    For aInx = LBound(aArray) To UBound(aArray): Do

        If aArray(aInx) = "" Then Exit Do   ' Next aInx
        bArray(bInx) = aArray(aInx)
        bInx = bInx + 1
        
    Loop While False: Next aInx

    '   If bArray is empty - Erase it
    '   Else - Compress bArray
    '
    If Not bInx > LBound(bArray) Then
        Erase bArray
    Else
        ReDim Preserve bArray(LBound(bArray) To bInx - 1)
    End If

    Misc_ArrayCompress = bArray

End Function

'   Is a single dimension Array empty? (Allocated but has no values)
'
'   https://forum.ozgrid.com/forum/index.php?thread/69797-determine-if-an-array-is-empty/
'
Public Function Misc_ArrayIsEmpty(ByRef aArray As Variant) As Boolean

    Misc_ArrayIsEmpty = (Len(Join(aArray, "")) = 0)
    
End Function

'   Is a single dimension Array Unallocated/Inaccessible?
'
Public Function Misc_ArrayIsNothing(ByRef aArray As Variant) As Boolean
Misc_ArrayIsNothing = True

    Dim Dummy As Long
    On Error Resume Next
        Dummy = Len(aArray(LBound(aArray)))
        Select Case Err.Number
            Case glbError_None
                ' Continue
            Case glbError_SubscriptOutOfRange
                Err.Clear
                Exit Function
            Case Else
                Stop: Exit Function
        End Select
    On Error GoTo 0
    
Misc_ArrayIsNothing = False
End Function

' =====================================================================
'   Registry
' =====================================================================

'   Does a Registry Key (ends with \) or Value name exist?
'
Public Function Misc_RegExist(ByVal RegName As String) As Boolean
Misc_RegExist = False

    Dim RegValue As Variant
    On Error Resume Next
        RegValue = glbWshShell.RegRead(RegName)
        Select Case Err.Number
            Case glbError_None
                '   Continue
            Case glbError_RegReadKeyNotFound
                Exit Function
            Case Else
                Stop: Exit Function
        End Select
    On Error GoTo 0

Misc_RegExist = True
End Function

' =====================================================================
'   Calendar
' =====================================================================

'   Open the associated Meeting for an Invite
'
'
Public Sub Misc_CalendarOpenMeetingFromInvite()
Const ThisProc = "CalendarOpenMeetingFromInvite"

    Dim InspItem As Object
    Dim ApptItem As AppointmentItem
        
    ' Get the Active Inspector Item (if any)
    '
    If Not Misc_GetActiveItem(InspItem, InspectorOnly:=True) Then Exit Sub
    
    ' Active Item must be an Invite (MeetingItem)
    '
    If Not (TypeOf InspItem Is Outlook.MeetingItem) Then
        Msg_Box Text:="Active Item is not an Invite (MeetingItem)", Proc:=ThisProc
        Exit Sub
    End If
    
    ' Associated Meeting (AppointmentItem) must exist
    '
    '    (The argument for GetAssociatedAppointment = False so that the appointment is not automatically added to the Calendar)
    '
    On Error Resume Next
        Set ApptItem = InspItem.GetAssociatedAppointment(False)
        If Err.Number = glbError_ItemIsZombie Then
            Msg_Box Proc:=ThisProc, Step:="GetAssociatedAppointment", Text:="The appointment/meeting for this invite has been moved, deleted, or declined."
            Exit Sub
        ElseIf Err.Number <> glbError_None Then
            Stop: Exit Sub
        End If
    On Error GoTo 0

    If ApptItem Is Nothing Then
        Msg_Box Proc:=ThisProc, Step:="ApptItem Is Nothing", Text:="The appointment/meeting for this invite has been moved, deleted, or declined."
        Exit Sub
    End If
    
    ' Open the assocaited Meeting/Appointment from the Calendar
    '
    On Error Resume Next
        ApptItem.GetInspector.Activate
            If Err.Number = glbError_OpenRecurringFromInstance Then
                Msg_Box Proc:=ThisProc, Step:="ApptItem.Display", Text:="Trying to show a Recurring Meeting from a Single Instance Update?"
                Exit Sub
            ElseIf Err.Number <> glbError_None Then
                Stop: Exit Sub
            End If
    On Error GoTo 0
    
End Sub

'   Has Invite for an Appointment been sent?
'
Public Function Misc_CalendarInviteSent(ByVal ApptItem As Outlook.AppointmentItem) As Boolean

    If Not Misc_OLGetProperty(ApptItem, glbPropTag_InviteSent, Misc_CalendarInviteSent) Then Exit Function
    
End Function

' =====================================================================
'   Reminders
' =====================================================================

Public Sub Misc_ReminderSnoozeBeforeStart_OneMinute()

    Misc_ReminderSnoozeBeforeStart (1)

End Sub
Public Sub Misc_ReminderSnoozeBeforeStart_OneHour()

    Misc_ReminderSnoozeBeforeStart (60)
    
End Sub
Private Sub Misc_ReminderSnoozeBeforeStart(ByVal Minutes As Long)
Const ThisProc = "Msc_ReminderSnoozeBeforeStart"

    Dim objRems As Outlook.Reminders
    Dim objRem As Outlook.Reminder
    Dim varTime As Variant

    Dim ItemsVisable As Integer
    ItemsVisable = 0
    
    Dim SnoozedList As String
    SnoozedList = ""

    Set objRems = Reminders
 
    For Each objRem In objRems
        If objRem.IsVisible = True Then
            ItemsVisable = ItemsVisable + 1
            varTime = DateDiff("n", Now(), objRem.Item.Start) - Minutes
            If varTime > 0 Then
              objRem.Snooze (varTime)
              SnoozedList = SnoozedList & "Snoozed " & CStr(varTime) & " min - '" & Left(objRem.Item.Subject, 32) & "'" & glbBlankLine
            Else
              SnoozedList = SnoozedList & "Too Late - '" & Left(objRem.Item.Subject, 32) & "'" & glbBlankLine
            End If
        End If
    Next objRem
    
    If ItemsVisable = 0 Then
        Msg_Box Text:="No visable Reminders.", Icon:=vbInformation, Proc:=ThisProc
    Else
        Msg_Box Text:=SnoozedList & glbBlankLine & "Snooze To " & Minutes & " min Before Start", Icon:=vbInformation
    End If

End Sub
