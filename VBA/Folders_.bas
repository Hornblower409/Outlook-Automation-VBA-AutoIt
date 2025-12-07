Attribute VB_Name = "Folders_"
Option Explicit
Option Private Module

' =====================================================================
'   IPF - Outlook Folder Types (ContainerClass)
'
'   Using Property:
'
'       PidTagContainerClass
'       Property ID: 0x3613
'       PR_CONTAINER_CLASS
'       ptagContainerClass
'       http://schemas.microsoft.com/mapi/proptag/0x3613001E
'       The value of this property MUST begin with "IPF.".
'
'   https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/folders-and-items-in-ews-in-exchange
'
'       "The folder class value is extensible. This means that the default FolderClass values are treated as prefixes
'       and you can add custom values. For example, you can create a folder with a FolderClass value of IPF.Contact.Contoso,
'       and it is treated as a Contacts folder".
'
'   https://github.com/libyal/libfmapi/blob/main/documentation/MAPI%20definitions.asciidoc#container_class_definitions
'
'       There is no official Outlook Type enumeration (that I can find). Only "known" text values:
'
'       IPF.Note                    Mail and Post
'       IPF.Note.OutlookHomepage    RSS Feeds
'       IPF.Appointment             Appointment/Meeting
'       IPF.Contact                 Contacts
'       IPF.Journal                 Journal entries
'       IPF.StickyNote              Post-It Notes
'       IPF.Task Folder             Tasks
'
'       To which I've added (thru necessity and digging around):
'
'       IPF.Configuration           Quick Steps and Conversation Actions
'       IPF.Imap                    IMAP Mail
'
' =====================================================================

    Private Const Folders_TypeAppointment   As String = "IPF.Appointment"
    Private Const Folders_TypeConfig        As String = "IPF.Configuration"
    Private Const Folders_TypeContact       As String = "IPF.Contact"
    Private Const Folders_TypeHomepage      As String = "IPF.Note.OutlookHomepage"
    Private Const Folders_TypeIMAP          As String = "IPF.Imap"
    Private Const Folders_TypeJournal       As String = "IPF.Journal"
    Private Const Folders_TypeMailPost      As String = "IPF.Note"
    Private Const Folders_TypeNote          As String = "IPF.StickyNote"
    Private Const Folders_TypeTask          As String = "IPF.Task"
    '
    Private Const Folders_TypeNoProp        As String = "IPF_NoProp"                    '   Folder does not have a glbPropTag_FolderType
    Private Const Folders_TypeUnknown       As String = "IPF_Unknown"                   '   Folder has an IPFType I've never seen before
    '
    Private Const Folders_TypeTable As String = _
        "   " & _
        "   Type                            |   Description                             " & vbLf & _
        "   " & _
            Folders_TypeAppointment & "     |   Calendar                                " & vbLf & _
            Folders_TypeConfig & "          |   Quick Steps and Conversation Actions    " & vbLf & _
            Folders_TypeContact & "         |   Contacts                                " & vbLf & _
            Folders_TypeHomepage & "        |   Outlook Today                           " & vbLf & _
            Folders_TypeIMAP & "            |   IMAP                                    " & vbLf & _
            Folders_TypeJournal & "         |   Journal                                 " & vbLf & _
            Folders_TypeMailPost & "        |   Mail and Post                           " & vbLf & _
            Folders_TypeNote & "            |   Sticky Note                             " & vbLf & _
            Folders_TypeTask & "            |   Task                                    " & vbLf & _
        "   " & _
            Folders_TypeNoProp & "          |   No IPF Property                         " & vbLf & _
            Folders_TypeUnknown & "         |   Unknown IPF Property Value              "
        '
        Private Const Folders_TypeTableColType      As Long = 0
        Private Const Folders_TypeTableColDesc      As Long = 1
    '

' =====================================================================
'   Known Paths
'
'       Any changes here must be reflected in Globals, Known Paths.
'
'       You can NOT keep a reference to a Folder object in a Global a long
'       time. Not sure why. I used to keep the reference to Known Paths but
'       after awhile the reference goes invalid.
'
'       Home Inbox is not a real thing. Just my choice for a folder that is usually
'       open in an Explorer and where I can go when I need a landing spot. Also used
'       by ProcXeq UpdateFolder scheduled task.
'
' =====================================================================

    Private Const Folders_KnownPathsTable As String = _
        "   " & _
        "   Path                                |   SessionFolder   |   Description                                         " & vbLf & _
        "   " & _
            glbKnownPath_Projects & "           |                   |   Projects - Current Projects                         " & vbLf & _
            glbKnownPath_Inbox & "              |                   |   Inbox - Default Inbox                               " & vbLf & _
            glbKnownPath_Drafts & "             |                   |   Drafts - Default Drafts                             " & vbLf & _
            glbKnownPath_GMailInbox & "         |   GMail           |   GMailInbox - xxxx@gmail.com Inbox           " & vbLf & _
            glbKnownPath_Deleted & "            |                   |   Deleted - Default Deleted Items                     " & vbLf & _
            glbKnownPath_Journal & "            |                   |   Journal - Default Journal                           " & vbLf & _
            glbKnownPath_Contacts & "           |                   |   Contacts - Default Contacts                         " & vbLf & _
            glbKnownPath_ProjectsArchive & "    |                   |   Archive - Current Projects Archive                  " & vbLf & _
            glbKnownPath_Outbox & "             |                   |   Outbox - Default Outbox                             "
        '
        Private Const Folders_KnownPathsTableColPath             As Long = 0
        Private Const Folders_KnownPathsTableColSessionFolder    As Long = 1
        Private Const Folders_KnownPathsTableColDescription      As Long = 2
    '
    
' =====================================================================
'   Known Paths
' =====================================================================

'   Check that all Known Paths exist.
'
'       Called from Globals_Init
'
Public Function Folders_KnownPathsCheck() As Boolean
Const ThisProc = "Folders_KnownPathsCheck"
Folders_KnownPathsCheck = False

    '   Get all Known Paths
    '
    Dim Paths() As String
    Paths = Folders_KnownPaths()
    
    '   Run thru the list and make sure all paths exist
    '
    Dim AnyError As Boolean
    Dim RowIndex As Long
    For RowIndex = 1 To UBound(Paths)
    
        Dim FolderPath As String
        FolderPath = Paths(RowIndex)
        
        If Folders_Path(FolderPath) Is Nothing Then
            AnyError = True
            Msg_Box Proc:=ThisProc, Step:="Folders_Path(KnownPath)", _
                    Text:="Outlook KnownPath does not exist. KnownPath: '" & FolderPath & "'."
        End If
        
    Next RowIndex
    If AnyError Then Exit Function

Folders_KnownPathsCheck = True
End Function

'   Get a KnownPath Folder
'
Public Function Folders_KnownPath(ByVal FolderPath As String) As Outlook.Folder

    If Not Tbl_TableConstExist(Folders_KnownPathsTable, FolderPath) Then
        Debug.Print "Path is not in the Folders_KnownPathsTable": Stop: Exit Function
    End If

    Set Folders_KnownPath = Folders_Path(FolderPath)
    If Not (Folders_KnownPath Is Nothing) Then Exit Function
    
    '   KnownPath FolderPath did not resolve to a Folder
    '
    '       Theoretically this can never happen. Globals_Init checked that all Known Paths exist.
    '       But with Stupid you never know.
    
    ' !!    Do NOT F5 this Stop.
    ' !!
    ' !!    You will be returning Nothing for a KnownPath Folder.
    ' !!    Walk back in the Stack and figure out what went wrong.
    ' !!
    '
    Stop: Stop: Stop

End Function

'   Return a Base One Array of all KnownPaths
'
Public Function Folders_KnownPaths() As String()

    '   Get the KnownPaths Table
    '
    Dim TableArray() As String
    TableArray = Tbl_TableConst(Folders_KnownPathsTable)

    Dim Paths() As String
    ReDim Paths(1 To UBound(TableArray, 1))
    
    Dim RowIndex As Long
    For RowIndex = 1 To UBound(TableArray, 1)
        Paths(RowIndex) = TableArray(RowIndex, Folders_KnownPathsTableColPath)
    Next RowIndex
    
    Folders_KnownPaths = Paths

End Function

' =====================================================================
'   IPF - Outlook Folder Types (ContainerClass)
' =====================================================================

'   Get the IPFType of a Folder
'
'   <-  Property Value. If it does not have the Property - Folders_TypeNoProp
'
Public Function Folders_Type(ByVal oFolder As Outlook.Folder) As String

    '   Get the Prop Value
    '
    If Not Misc_OLGetProperty(oFolder, glbPropTag_IPFFolderType, Folders_Type) Then
        Folders_Type = Folders_TypeNoProp
    End If

End Function

'   Get the IPFType Description of a Folder
'
'   <-  Type Description from my table. If not found - Folders_TypeUnknown Description
'
Public Function Folders_TypeDesc(ByVal oFolder As Outlook.Folder) As String

    Dim IPFType As String
    IPFType = Folders_Type(oFolder)
    
    Dim TypeDesc As String
    If Not Tbl_TableConstFind(Folders_TypeTable, IPFType, Folders_TypeTableColDesc, TypeDesc) Then
        If Not Tbl_TableConstFind(Folders_TypeTable, Folders_TypeUnknown, Folders_TypeTableColDesc, TypeDesc) Then Stop: Exit Function
    End If
    
    Folders_TypeDesc = TypeDesc

End Function

'   Is a Folder an IPFType IMAP?
'
Public Function Folders_TypeIsIMAP(ByVal oFolder As Outlook.Folder) As Boolean

    Folders_TypeIsIMAP = (Folders_Type(oFolder) = Folders_TypeIMAP)

End Function

' =====================================================================
'   Build a Dict of all the Folders for all the Stores in this Session
' =====================================================================
'
'   Key = FolderPath, Item = Folder Object
'
'   Calling:
'
'       Dim dAllFolders As Scripting.Dictionary
'       Set dAllFolders = New Scripting.Dictionary
'       If Not Folders_AllFolders(dAllFolders) Then Stop: Exit Function
'
'   Using:
'
'       Dim vFolder As Variant
'       Dim oFolder As Outlook.Folder
'       For Each vFolder In dAllFolders.Items
'           Set oFolder = vFolder
'           ...
'           ...
'           ...
'       Next vFolder
'
Public Function Folders_AllFolders(ByVal dAllFolders As Scripting.Dictionary) As Boolean
Const ThisProc = "Folders_AllFolders"
Folders_AllFolders = False

    On Error GoTo Error_Exit
    
    '   Walk all the Stores for Normal Folders
    '   Start off Folders_AllFoldersDecend from the Store Root Folder
    '
    Dim oStore As Outlook.Store
    For Each oStore In Session.Stores
        If Not Folders_AllFoldersDecend(oStore.GetRootFolder, dAllFolders) Then Stop: Exit Function
    Next oStore
    
    '   Walk all the Stores for Search Folders
    '   Search Folders have no Subfolders. No Decend.
    '
    Dim oSearchFolders As Outlook.Folders
    Dim oSearchFolder As Outlook.Folder
    For Each oStore In Session.Stores
        Set oSearchFolders = oStore.GetSearchFolders
        For Each oSearchFolder In oSearchFolders
            dAllFolders.Add oSearchFolder.FolderPath, oSearchFolder
        Next oSearchFolder
    Next oStore

    Folders_AllFolders = True
    Exit Function

Error_Exit:
Stop: Exit Function
End Function

'   Add to a Dict a Folder and any of it's Subfolders
'   Called from Folders_AllFolders and calls itself recursive.
'
'   Key = FolderPath, Item = Folder Object
'
Private Function Folders_AllFoldersDecend(ByVal oFolder As Outlook.Folder, ByVal dAllFolders As Scripting.Dictionary) As Boolean
Folders_AllFoldersDecend = False

    '   Add the current Folder
    '
    dAllFolders.Add oFolder.FolderPath, oFolder

    '   For any subfolders - recursive call
    '
    Dim oSubFolder As Outlook.Folder
    For Each oSubFolder In oFolder.Folders
        If Not Folders_AllFoldersDecend(oSubFolder, dAllFolders) Then Stop: Exit Function
    Next oSubFolder

Folders_AllFoldersDecend = True
End Function

' =====================================================================
'   Get a Folder for {Something}
' =====================================================================

'   Get an Item's Folder
'
'   If found - Returns TRUE and oFolder set.
'   Else - Returns FALSE and oFolder = Nothing.
'
Public Function Folders_Item(ByVal oItem As Object, ByRef oFolder As Outlook.Folder) As Boolean
Folders_Item = False
Set oFolder = Nothing

    '   Never seen a Depth > 3. If we get to 5 - something is hoarked.
    '
    Dim oItemParent As Object
    Set oItemParent = oItem
    Dim Depth As Long
    For Depth = 1 To 5

        On Error GoTo Error_Exit
            Set oItemParent = oItemParent.Parent
            If TypeOf oItemParent Is Outlook.Folder Then Set oFolder = oItemParent
        On Error GoTo 0
        If Not oFolder Is Nothing Then
            Folders_Item = True
            Exit Function
        End If
            
    Next Depth

Error_Exit:
End Function

'   Get the Deleted Items Folder for an Item
'
Public Function Folders_Deleted(ByVal oItem As Object, ByRef oDeletedFolder As Outlook.Folder) As Boolean
Folders_Deleted = False

    Dim oFolder As Outlook.Folder
    If Not Folders_Item(oItem, oFolder) Then Exit Function
    
    On Error GoTo Error_Exit
        Set oDeletedFolder = oFolder.Store.GetDefaultFolder(Outlook.olFolderDeletedItems)
    On Error GoTo 0
    
    Folders_Deleted = True
    Exit Function

Error_Exit:
End Function

'   Get a Window's Folder
'
Public Function Folders_Window(ByVal Window As Object, ByRef Folder As Outlook.Folder) As Boolean
Folders_Window = False

    If TypeOf Window Is Outlook.Inspector Then
        If Not Folders_Item(Window.CurrentItem, Folder) Then Stop: Exit Function
    ElseIf TypeOf Window Is Outlook.Explorer Then
        Set Folder = Window.CurrentFolder
    Else
        Stop: Exit Function
    End If

Folders_Window = True
End Function

'   Get a Folder from a FolderPath
'
Public Function Folders_Path(ByVal FolderPath As String) As Outlook.Folder
Set Folders_Path = Nothing

    '   If it's a KnownPath and it has a SessionFolder name in the table
    '   - Use the SessionFolder method to get the Folder
    '
    Dim SessionFolder As String
    If Tbl_TableConstFind(Folders_KnownPathsTable, FolderPath, Folders_KnownPathsTableColSessionFolder, SessionFolder) Then
        If SessionFolder <> "" Then
            Set Folders_Path = IMAP_SessionInbox(FolderPath, SessionFolder)
            Exit Function
        End If
    End If

    '   If it doesn't start with "\\" it's not a Full Folder Path
    '
    If Left(FolderPath, 2) <> "\\" Then
        Debug.Print "Invalid FolderPath: " & FolderPath
        Stop: Exit Function
    End If
    
    '   Strip the leading "\\" so it matches the Store Root Folder oFolder.Name
    '   Split the path into an Array on "\"
    '
    Dim PathParts As Variant
    PathParts = Split(Right(FolderPath, Len(FolderPath) - 2), "\")
    
    '   URL Percent Decode ("\/%") each part of the path
    '
    '       SPOS - Stupid URL encodes "\/%" in oFolder.FolderPath, but not in oFolder.Name
    '
    '       He has to encode in FolderPath, but he could at least be consistent. Assume this
    '       is so the user, who only sees oFolder.Name in the Navigation Pane, won't be confused.
    '
    '       Which leads to the question - Why allow "\/" in a Folder Name in the first place?
    '       For some kind of IMAP account that doesn't recognize them as path seperators?
    '
    '       AND THEN - He also allows other junk {Tabs that I know of) in oFolder.Name
    '
    '   The base for my code was "Get a Folder Object from a Folder Path"
    '   https://learn.microsoft.com/en-us/office/vba/outlook/how-to/items-folders-and-stores/obtain-a-folder-object-from-a-folder-path
    '
    '   But it didn't handle encoding. I had to add it.
    '
    Dim PartsIx As Long
    For PartsIx = 0 To UBound(PathParts)
        PathParts(PartsIx) = Misc_URLDecode(PathParts(PartsIx))
    Next PartsIx

    '   oFolder = Store Root Folder
    '
    Dim oFolder As Outlook.Folder
    On Error Resume Next

        '   Just Set and trap instead of walking the collection
        '
        Set oFolder = Session.Folders.Item(PathParts(0))
        Select Case Err.Number
            Case glbError_None
            Case glbError_ObjectNotFound
                Exit Function
            Case Else
                Stop: Exit Function
        End Select

    On Error GoTo 0
    If oFolder Is Nothing Then Exit Function

    '   For each part of the path after the first (Store)
    '
    '       Look thru the subfolders.
    '       Falls out of the For with oFolder = the one I want
    '       or Exits if any part not found.
    '
    Dim SubFolders As Outlook.Folders
    For PartsIx = 1 To UBound(PathParts)

        '   If this Part doesn't support SubFolders - error exit
        '
        On Error Resume Next
            Set SubFolders = oFolder.Folders
            If Err.Number <> glbError_None Then Exit Function
        On Error GoTo 0
        
        '   Check this Part
        '
        On Error Resume Next
        
            Set oFolder = SubFolders.Item(PathParts(PartsIx))
            Select Case Err.Number
                Case glbError_None
                Case glbError_ObjectNotFound
                    Exit Function
                Case Else
                    Stop: Exit Function
            End Select

        On Error GoTo 0
        If oFolder Is Nothing Then Exit Function

    Next PartsIx

    '   Check for an IMAP Folder that did not use the SessionFolder method
    '
    If Folders_TypeIsIMAP(oFolder) Then
        Debug.Print "Folder is an IMAP folder but did not use the SessionFolder method. FolderPath: " & FolderPath
        Stop: Exit Function
    End If
    
    Set Folders_Path = oFolder
    
End Function

' =====================================================================
'   Get {Something} for a Folder
' =====================================================================

'   Get an open Explorer for a folder (if any)
'
Public Function Folders_FolderExplorer(ByVal Folder As Outlook.Folder) As Outlook.Explorer

    Set Folders_FolderExplorer = Nothing

    ' Search Explorers for one that is open to this folder
    '
    Dim Explorer As Outlook.Explorer
    For Each Explorer In Application.Explorers
    
        '   Got it - done
        '
        If Explorer.CurrentFolder.FolderPath = Folder.FolderPath Then
            Set Folders_FolderExplorer = Explorer
            Exit Function
        End If
        
    Next
    
End Function
