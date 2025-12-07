Attribute VB_Name = "File_"
Option Explicit
Option Private Module

' =====================================================================
'   Read/Write/Append a String From/To a File
' =====================================================================

'   Read an entire text file into a string.
'
Public Function File_ReadText(ByVal FileSpec As String, ByRef InputText As String) As Boolean
Const ThisProc = "File_ReadText"
File_ReadText = False
    
    '   Open a TextStream to the file and get the whole thing
    '
    Dim oFile As Scripting.TextStream
    On Error Resume Next

        Set oFile = glbFSO.OpenTextFile(FileSpec, IOMode:=ForReading, Create:=False, Format:=TristateTrue)
        If Err.Number <> glbError_None Then
            Msg_Box oErr:=Err, Proc:=ThisProc, Step:="OpenTextFile", Text:="Open text file failed." & vbNewLine & "File = '" & FileSpec & "'."
            Exit Function
        End If
        
        InputText = oFile.ReadAll()
        If Err.Number <> glbError_None Then
            Msg_Box oErr:=Err, Icon:=vbCritical, Proc:=ThisProc, Step:="ReadAll", Text:=FileSpec
            Msg_Box oErr:=Err, Proc:=ThisProc, Step:="TextStream.ReadAll", Text:="Read text file failed." & vbNewLine & "File = '" & FileSpec & "'."
            Exit Function
        End If
        
        oFile.Close
        If Err.Number <> glbError_None Then
            Msg_Box oErr:=Err, Proc:=ThisProc, Step:="TextStream.Close", Text:="Close text file failed." & vbNewLine & "File = '" & FileSpec & "'."
            Exit Function
        End If
        
    On Error GoTo 0

File_ReadText = True
End Function

'   Write a string as a Text File (Overwrites any existing file)
'
Public Function File_WriteText(ByVal FileSpec As String, ByVal TextString As String) As Boolean
Const ThisProc = "File_WriteText"
File_WriteText = False

    On Error Resume Next
    
        Dim TextStream As Scripting.TextStream
        Set TextStream = glbFSO.CreateTextFile(FileSpec, Overwrite:=True, Unicode:=True)
        If Err.Number <> glbError_None Then
            Msg_Box oErr:=Err, Proc:=ThisProc, Step:="FSO.CreateTextFile", Text:="Create text file failed." & vbNewLine & "File = '" & FileSpec & "'."
            Exit Function
        End If
    
        TextStream.Write TextString
        If Err.Number <> glbError_None Then
            Msg_Box oErr:=Err, Proc:=ThisProc, Step:="TextStream.Write", Text:="Write text file failed." & vbNewLine & "File = '" & FileSpec & "'."
            Exit Function
        End If
        
        TextStream.Close
        If Err.Number <> glbError_None Then
            Msg_Box oErr:=Err, Proc:=ThisProc, Step:="TextStream.Close", Text:="Close text file failed." & vbNewLine & "File = '" & FileSpec & "'."
            Exit Function
        End If
    
    On Error GoTo 0
    
File_WriteText = True
End Function

'   Appends a string + vbNewLine to a Text File (Creates the file if it doesn't exist)
'
Public Function File_AppendText(ByVal FileSpec As String, ByVal TextString As String) As Boolean
Const ThisProc = "File_AppendText"
File_AppendText = False

    On Error Resume Next
    
        Dim TextStream As Scripting.TextStream
        Set TextStream = glbFSO.OpenTextFile(FileSpec, IOMode:=ForAppending, Create:=True, Format:=TristateUseDefault)
        If Err.Number <> glbError_None Then
            Msg_Box oErr:=Err, Proc:=ThisProc, Step:="FSO.OpenTextFile", Text:="Open text file failed." & vbNewLine & "File = '" & FileSpec & "'."
            Exit Function
        End If
    
        TextStream.WriteLine TextString
        If Err.Number <> glbError_None Then
            Msg_Box oErr:=Err, Proc:=ThisProc, Step:="TextStream.WriteLine", Text:="WriteLine text file failed." & vbNewLine & "File = '" & FileSpec & "'."
            Exit Function
        End If
        
        TextStream.Close
        If Err.Number <> glbError_None Then
            Msg_Box oErr:=Err, Proc:=ThisProc, Step:="TextStream.Close", Text:="Close text file failed." & vbNewLine & "File = '" & FileSpec & "'."
            Exit Function
        End If
    
    On Error GoTo 0

File_AppendText = True
End Function

' =====================================================================
'   MkDir
' =====================================================================

'   SPOS - "glbFSO.CreateFolder FolderPath" is suppose to create all intermediate sub folders.
'   But it fails (randomly?). Advice from the web is to create each subfolder individually.
'
'   From https://stackoverflow.com/a/76318929/28691646
'
'   Recursive walk from the bottom up.
'
Public Function File_MkDir(ByVal FolderPath As String) As Boolean
File_MkDir = False

    On Error GoTo Error_Exit
    
        If Not glbFSO.FolderExists(glbFSO.GetParentFolderName(FolderPath)) Then File_MkDir glbFSO.GetParentFolderName(FolderPath)
        If Not glbFSO.FolderExists(FolderPath) Then glbFSO.CreateFolder FolderPath

    On Error GoTo 0
    
    File_MkDir = True
    Exit Function
    
Error_Exit:
    Stop: Exit Function
End Function

' =====================================================================
'   Backup VbaProject.OTM
' =====================================================================

'   Make a Manual backup of the VbaProject.OTM file and delete older versions.
'
Public Function File_BackupVBAProject_Manual() As Boolean
Const ThisProc = "File_BackupVBAProject_Manual"
File_BackupVBAProject_Manual = False

    '   SPOS - InputBox is a fixed width and only wraps text on a space. No spaces
    '   in a long string and he just truncates.
    '
    Dim Prefix As String
    Prefix = VBA.InputBox(Title:=ThisProc, Default:="Manual", _
                      Prompt:="Folder: '" & glbVBAProjectBackupFolder & "'." & vbNewLine & _
                              "File: '{Prefix}_DateTime_VbaProject.OTM'." & glbBlankLine & _
                              "Will be deleted when there are more than " & glbVBAProjectBackupKeepers & " versions." & glbBlankLine & _
                              "You MUST have done a 'File -> Save VBAProject.OTM' for this to work." & glbBlankLine & _
                              "File Name Prefix (Blank to Cancel).")
    If Prefix = "" Then
        Msg_Box Proc:=ThisProc, Step:="File Name Prefix Check", Text:="Backup CANCELED."
        Exit Function
    End If
        
    If Not File_BackupVBAProject_Trim(Prefix:=Prefix) Then Exit Function

File_BackupVBAProject_Manual = True
End Function

'   Backup the VbaProject.OTM file and Delete Older Versions.
'
'   Called from Application_Startup, Application_Init_Link (on the Ribbon), and Manual Backup
'
Public Function File_BackupVBAProject_Trim(Optional ByVal Prefix As String = "") As Boolean
Const ThisProc = "File_BackupVBAProject_Trim"
File_BackupVBAProject_Trim = False

    '   Do the backup and delete older versions
    '
    If Not File_BackupVBAProject(glbVBAProjectBackupFolder, Prefix:=Prefix) Then Exit Function
    If Not File_DeleteOldByCreated(glbVBAProjectBackupFolder, glbVBAProjectBackupKeepers) Then Exit Function

File_BackupVBAProject_Trim = True
End Function

'   Backup the VbaProject.OTM file
'
Private Function File_BackupVBAProject(ByVal BackupFolder As String, Optional ByVal Prefix As String = "") As Boolean
Const ThisProc = "File_BackupVBAProject"
File_BackupVBAProject = False
    
    On Error GoTo ErrorExit

    '   Build the full path to VbaProject.OTM
    '
    Dim Source As String
    If Not Misc_EnvironmentGet("APPDATA", Source) Then Exit Function
    Source = Source & "\Microsoft\Outlook\VbaProject.OTM"
    
    '   Build File Name Prefix
    '
    Dim FileNamePrefix As String
    If Prefix <> "" Then FileNamePrefix = Trim(Prefix) & "_"
    
    '   Build full path to the backup file
    '
    Dim Destination As String
    Destination = BackupFolder & "\" & FileNamePrefix & Misc_NowStamp() & "_VbaProject.OTM"
    
    '   Files and Folders checks
    '
    If Not glbFSO.FileExists(Source) Then
        Msg_Box Proc:=ThisProc, Step:="File Exist?", Text:="Source file '" & Source & "' does not exist."
        Exit Function
    End If
    
    If Not glbFSO.FolderExists(BackupFolder) Then
        Msg_Box Proc:=ThisProc, Step:="Folder Exist?", Text:="Backup folder '" & BackupFolder & "' does not exist."
        Exit Function
    End If
    
    If glbFSO.FileExists(Destination) Then
        Msg_Box Proc:=ThisProc, Step:="File Exist?", Text:="Backup file '" & Destination & "' already exist."
        Exit Function
    End If
    
    '   Do the Copy
    '
    glbFSO.CopyFile Source, Destination
    
    '   Exit with OK
    '
    File_BackupVBAProject = True
    Exit Function

ErrorExit:

    Msg_Box oErr:=Err, Proc:=ThisProc

End Function

'   Delete older files in a folder based on DateCreated (NOT DateLastModified)
'
'   Called by File_BackupVBAProjectTrim
'
'   SPOS - Because VbaProject.OTM does not change it's DateLastModified or size on a reliable basis
'   we have to look for older files based on the DateCreated (when we made the copy) not DateLastModified of the original.
'   See: https://stackoverflow.com/questions/24816147/outlook-vbaproject-otm-timestamp-is-not-updated-upon-changing
'
Public Function File_DeleteOldByCreated(ByVal FolderPath As String, ByVal VersionsToKeep As Long) As Boolean
Const ThisProc = "File_DeleteOldByCreated"
File_DeleteOldByCreated = False

    On Error GoTo ErrorExit
    
    If Not glbFSO.FolderExists(FolderPath) Then
        Msg_Box Proc:=ThisProc, Step:="Folder Exist?", Text:="Folder '" & FolderPath & "' does not exist."
        Exit Function
    End If
    
    Dim oFolder As Scripting.Folder
    Set oFolder = glbFSO.GetFolder(FolderPath)
    
    Dim oFiles As Scripting.Files
    Set oFiles = oFolder.Files
    
    Do While oFiles.count > VersionsToKeep
    
        '   Get the DT Created of the first file in the Collection
        '
        Dim oOldest As Scripting.File

            ' SPOS - You can't get a member of the Files collection by index number.
            '
            '   https://stackoverflow.com/questions/848851/asp-filesystemobject-collection-cannot-be-accessed-by-index
            '
            '   "In general, collections can be accessed via index numbering, but the Files Collection is not a normal collection.
            '   It does have an item property, but it appears that the key that it uses is filename"
            '
            ' Set oOldest = oFiles.Item(1)  --> BOOM
            '
            ' So we do a For Each, get the first file, and then immediatley exit the For.
            '
            Dim oFile As Scripting.File
            For Each oFile In oFiles
                Set oOldest = oFile
                Exit For
            Next oFile
        
        '   Find the oldest file in the collection and Delete it
        '
        For Each oFile In oFiles
            If oFile.DateCreated < oOldest.DateCreated Then Set oOldest = oFile
        Next oFile
        glbFSO.DeleteFile oOldest.Path
    
    Loop
    
    '   Exit OK
    '
    File_DeleteOldByCreated = True
    Exit Function

ErrorExit:

    Msg_Box oErr:=Err, Proc:=ThisProc

End Function

' =====================================================================
'   Save HTML in a File
' =====================================================================

Public Function File_SaveInspectorHTML() As Boolean
File_SaveInspectorHTML = False

    Dim Item As Object
    If Not Misc_GetActiveItem(Item, InspectorOnly:=True) Then Exit Function
    If Not Mail_HasHTMLBody(Item) Then Stop: Exit Function
    
    Dim Prefix As String
    Prefix = VBA.InputBox("File Name Prefix  (Optional)")
    If Not File_SaveHTML(Item.HTMLBody, Prefix) Then Stop: Exit Function

File_SaveInspectorHTML = True
End Function

Public Function File_SaveHTML(ByVal sHTMLBody As String, Optional ByVal FilePrefix As String = "") As Boolean
File_SaveHTML = False

    Dim Prefix As String
    Prefix = FilePrefix
    
    '   Build full path to the file
    '
    Dim FolderPath As String: FolderPath = glbFileSaveHTML
    
    If Prefix <> "" Then Prefix = Trim(Prefix) & "_"
    Dim Destination As String
    Destination = FolderPath & "\" & Prefix & Misc_NowStamp() & "_Outlook_HTML.html"

    '   2024-10-29 Switch to File_WriteText
    '
    ' If Not File_FileWrite(Destination, sHTMLBody) Then Exit Function
    If Not File_WriteText(Destination, sHTMLBody) Then Exit Function

File_SaveHTML = True
End Function

' =====================================================================
'   File/Folder Picker
' =====================================================================

'   Open a Folder Select Dialog and get a Full Folder Path (e.g. "C:\XXX\")
'
'   SPOS - Outlook (unlike all other Office apps) doesn't support the Application.FileDialog object
'   so we have to use the App Shell VBA COM Object.
'
'   Options. Add together. (From https://learn.microsoft.com/en-us/windows/win32/api/shlobj_core/ns-shlobj_core-browseinfoa)
'
'       BIF_RETURNONLYFSDIRS    = &H1           Default. Only return normal directories.
'       BIF_RETURNFSANCESTORS   = &H8           {Not Working} Shows folders above the FolderStartPath.
'       BIF_EDITBOX             = &H10          Edit Box. User can type a folder name.
'       BIF_NEWDIALOGSTYLE      = &H40          {Not Working} Use new, larger, fancy dialog box.
'       BIF_NONEWFOLDERBUTTON   = &H200         Don't show the New Folder button.
'       BIF_BROWSEINCLUDEFILES  = &H4000        Dialog box displays files as well as folders.
'

'
Public Function File_FolderSelect( _
    ByRef FolderPath As String, _
    Optional ByVal Caption As String = "Select Folder", _
    Optional ByVal FolderStartPath As Variant, _
    Optional ByVal Options As Long = 1) _
    As Boolean
Const ThisProc = "File_FolderSelect"
File_FolderSelect = False
    
    Dim oFolder As Shell32.Folder
    
    Set oFolder = glbAppShell.BrowseForFolder(0, Caption, Options, FolderStartPath)
    If oFolder Is Nothing Then Exit Function
    
    '   2025-01-18 - RubberDuck objects to oFolder.Self.Path. But error without it.
    '
    '@Ignore MemberNotOnInterface
    FolderPath = oFolder.Self.Path
    
    If FolderPath = "" Then Exit Function
    FolderPath = FolderPath & "\"
    
File_FolderSelect = True
End Function

'   Get a valid existing File Folder Path and Overwrite switch if it's not empty
'
Public Function File_FolderSelectOverwrite( _
    ByRef FolderPath As String, _
    ByRef Overwrite As Boolean, _
    Optional ByVal Caption As String, _
    Optional ByVal FolderStartPath As Variant, _
    Optional ByVal Options As Long = 1) _
    As Boolean
Const ThisProc = "File_FolderSelectOverwrite"
File_FolderSelectOverwrite = False

    '   Loop until we get a valid folder
    '
    Do
    
        '   Get a folder path
        '
        Do
            If Not File_FolderSelect( _
                FolderPath:=FolderPath, _
                Caption:=Caption, _
                FolderStartPath:=FolderStartPath, _
                Options:=Options _
            ) Then Exit Function
            
            '   If the folder doesn't exist - ask to create it
            '
            If glbFSO.FolderExists(FolderPath) Then Exit Do
        
            Select Case Msg_Box(Text:="Folder '" & FolderPath & "' does not exist." & glbBlankLine & "Create it?", Icon:=vbQuestion, _
                        Buttons:=vbYesNo, Default:=vbDefaultButton1, Proc:=ThisProc, Step:="Create folder?")
                Case vbYes
                    '   2025-01-18 - Replaced glbFSO.CreateFolder with File_MkDir
                    '   glbFSO.CreateFolder FolderPath
                    If Not File_MkDir(FolderPath) Then Stop: Exit Function
                    Exit Do
                Case vbNo
                    ' Loop
            End Select
        
        Loop
            
        '   If the folder is not empty - warn about over-write
        '
        Dim oFileFolder As Scripting.Folder
        Set oFileFolder = glbFSO.GetFolder(FolderPath)
        Dim oSubFolders As Scripting.Folders
        Set oSubFolders = oFileFolder.SubFolders
        Dim oFiles As Scripting.Files
        Set oFiles = oFileFolder.Files
        If oFiles.count = 0 And oSubFolders.count = 0 Then Exit Do
        
        Select Case Msg_Box(Text:="Folder '" & FolderPath & "' is not empty." & glbBlankLine & "Overwrite any existing files in this folder?", Icon:=vbQuestion, _
                    Buttons:=vbYesNoCancel, Default:=vbDefaultButton2, Proc:=ThisProc, Step:="Check Folder Empty")
            Case vbYes
                Overwrite = True
                Exit Do
            Case vbNo
                Overwrite = False
                Exit Do
            Case vbCancel
                ' Continue
        End Select
        
    Loop

File_FolderSelectOverwrite = True
End Function

'   Open a Standard File Open Dialog via AutoIt
'
'       If you Option MULTISELECT then you always get back a Variant Array,
'       even if only one file was selected. If not MULTISELECT then you get
'       back a Variant String.
'
Public Function File_FileOpenDialog( _
    ByRef FileSpecs As Variant, _
    ByVal Caller As String, _
    Optional ByVal Title As String = "File Open", _
    Optional ByVal InitDir As String = "", _
    Optional ByVal Filter As String = "All (*.*)", _
    Optional ByVal Options As Long = 0, _
    Optional ByVal InitFile As String = "" _
    ) As Boolean
File_FileOpenDialog = False
    
'   Options (Add together)
'
'   $FD_FILEMUSTEXIST   (1) = File Must Exist (if user types a filename)
'   $FD_PATHMUSTEXIST   (2) = Path Must Exist (if user types a path, ending with a backslash)
'   $FD_MULTISELECT     (4) = Allow MultiSelect
'   $FD_PROMPTCREATENEW (8) = Prompt to Create New File (if does not exist)

    '   Build the Command Line
    '
    Dim CmdLine As String
    CmdLine = Join(Array(glbFileDialogs_Action_FileOpen, Title, InitDir, Filter, Options, InitFile), Chr(glbFileDialogs_CmdLineSep))
    
    '   Call FileDialogs
    '
    Dim Results As String
    If Not File_FileDialogs(Results, CmdLine, Caller) Then Exit Function
    
    '   If they didn't ask for a Multi - give them a Variant String
    '
    If (Options And 4) <> 4 Then
        FileSpecs = Results
        File_FileOpenDialog = True
        Exit Function
    End If
    
    '   They ask for a Multi
    '
    '       But only one file came back - give them a single element Variant Array
    '
    If InStr(Results, "|") = 0 Then
        ReDim FileSpecs(0)
        FileSpecs(0) = Results
        File_FileOpenDialog = True
        Exit Function
    End If

    '       Multiple files came back - Convert "Directory|file1|file2|..." into an Array of FileSpecs
    '
    Dim Directory As String
    Directory = Left(Results, InStr(Results, "|") - 1) & "\"
    Results = Mid(Results, InStr(Results, "|") + 1)
    
    FileSpecs = Split(Results, "|")
    Dim Inx As Long
    For Inx = 0 To UBound(FileSpecs)
        FileSpecs(Inx) = Directory & FileSpecs(Inx)
    Next Inx
    
File_FileOpenDialog = True
End Function

'   Open a Standard File Save Dialog via AutoIt
'
Public Function File_FileSaveDialog( _
    ByRef FileSpec As String, _
    ByVal Caller As String, _
    Optional ByVal Title As String = "File Save", _
    Optional ByVal InitDir As String = "", _
    Optional ByVal Filter As String = "All (*.*)", _
    Optional ByVal Options As Long = 0, _
    Optional ByVal DefaultName As String = "" _
    ) As Boolean
File_FileSaveDialog = False
    
'   Options (Add together)
'
'   $FD_PATHMUSTEXIST   (2)  = Path Must Exist (if user types a path, ending with a backslash)
'   $FD_PROMPTOVERWRITE (16) = Prompt to OverWrite File

    '   Build the Command Line
    '
    Dim CmdLine As String
    CmdLine = Join(Array(glbFileDialogs_Action_FileSave, Title, InitDir, Filter, Options, DefaultName), Chr(glbFileDialogs_CmdLineSep))
    
    '   Call FileDialogs
    '
    If Not File_FileDialogs(FileSpec, CmdLine, Caller) Then Exit Function
    
File_FileSaveDialog = True
End Function

'   Open a Standard Folder Select Dialog via AutoIt
'
'       No trailing Backslash on the Folder Path
'
Public Function File_FolderSelectDialog( _
    ByRef FolderPath As String, _
    ByVal Caller As String, _
    Optional ByVal Title As String = "File Save", _
    Optional ByVal RootDir As String = "", _
    Optional ByVal InitDir As String = "" _
    ) As Boolean
File_FolderSelectDialog = False
    
    '   Build the Command Line
    '
    Dim CmdLine As String
    CmdLine = Join(Array(glbFileDialogs_Action_FolderSelect, Title, RootDir, InitDir), Chr(glbFileDialogs_CmdLineSep))
    
    '   Call FileDialogs
    '
    If Not File_FileDialogs(FolderPath, CmdLine, Caller) Then Exit Function
    
File_FolderSelectDialog = True
End Function

'   Backend for all FileDialogs
'
Private Function File_FileDialogs(ByRef Results As String, ByVal CmdLine As String, ByVal Caller As String) As Boolean
File_FileDialogs = False

    glbWshShell.RegWrite glbFileDialogs_RegBaseKey, ""
    glbWshShell.RegWrite glbFileDialogs_RegCmdLine, CmdLine, "REG_SZ"
    
    glbWshShell.RegWrite glbFileDialogs_RegRetunKey, ""
    glbWshShell.RegWrite glbFileDialogs_RegResults, "", "REG_SZ"
    glbWshShell.RegWrite glbFileDialogs_RegCanceled, "99", "REG_SZ"
    
    Dim ExitCode As Integer
    ExitCode = Utility_ShellRun(glbHotRodLnks & glbHotRodLnks_FileDialogs, Wait:=True)
    
    '   Non-Zero Exit Code - AutoIt FileDialogs has already shown a message and done a hr_Error_Exit
    '
    If ExitCode <> 0 Then Exit Function
    
    Dim Canceled As String
    Canceled = glbWshShell.RegRead(glbFileDialogs_RegCanceled)
    Select Case Canceled
        Case "0"
            ' Continue
        Case "1"
            Exit Function
            
        '   Value is the same as before I called. Something went wrong.
        '
        Case "99"
            Msg_Box Proc:=Caller, Step:="AutoIt Call Check", _
                    Text:="AutoIt FileDialogs call Failed Completely."
            Exit Function
            
        Case Else
            Stop: Exit Function
    End Select
    
    Results = glbWshShell.RegRead(glbFileDialogs_RegResults)
    
File_FileDialogs = True
End Function

' =====================================================================
'   Temp File
' =====================================================================

'   Get a unique, non-existing Temp file spec
'
Public Function File_GetTempFileSpec(ByRef FileSpec As String) As Boolean
Const ThisProc = "File_GetTempFileSpec"
File_GetTempFileSpec = False
        
    Dim FileName As String
    Dim FileFound As Boolean
    Dim LoopIx As Long
    
    FileFound = True
    For LoopIx = 1 To 100
    
        FileName = Misc_NowStamp() & ".MSG"
        FileSpec = glbTempFilePath & FileName
        
        If Not glbFSO.FileExists(FileSpec) Then
            FileFound = False
            Exit For
        End If
        
        Sleep 100
    
    Next LoopIx
    
    If FileFound Then
        Msg_Box Proc:=ThisProc, Step:="Test generated File Spec", Text:="Could not generate a unique non-existing File Spec."
        Exit Function
    End If
    
    File_GetTempFileSpec = True

End Function

'   Save an Item to a Temp file
'
Public Function File_SaveToTemp(ByVal Item As Object, ByRef FileSpec As String) As Boolean
Const ThisProc = "File_SaveToTemp"
File_SaveToTemp = False

    '   Get a Temp File Spec
    '
    If Not File_GetTempFileSpec(FileSpec) Then Exit Function
    
    '   Save the Item to the file
    '
    If Item Is Nothing Then Stop: Exit Function
    Item.SaveAs FileSpec, Outlook.olMSGUnicode
    
    '   And make sure it's there
    '
    If Not glbFSO.FileExists(FileSpec) Then
        Msg_Box Proc:=ThisProc, Step:="After SaveAs File Exist?", Text:="File was not created '" & FileSpec & "'."
        Exit Function
    End If
    
File_SaveToTemp = True
End Function

'   Load an Item from a file into an Outlook Folder & delete the file
'
Public Function File_LoadFromFile(ByRef oItem As Object, ByVal FileSpec As String, ByVal oFolder As Outlook.Folder) As Boolean
Const ThisProc = "File_LoadFromFile"
File_LoadFromFile = False

    Set oItem = Nothing
    
    If Not glbFSO.FileExists(FileSpec) Then
        Msg_Box Proc:=ThisProc, Step:="Before Load file check", Text:="File does not exist '" & FileSpec & "'."
        Exit Function
    End If
    
    '   Load the Item - using OpenSharedItem
    '
    '   Creates an in-memory Item in the Default Inbox. But Stupid locks the file until I Move the
    '   Item to a different folder. So when the target folder IS the Default Inbox, I have to Move
    '   it someplace else (Deleted) and then back to the Default Inbox. So I can delete the file.
    '
    '   2025-03-11 - Tried CreateItemFromTemplate but it has it's own problems.
    
    Set oItem = Session.OpenSharedItem(FileSpec)
    If oItem Is Nothing Then Stop: Exit Function
    
    If oFolder.FolderPath = glbKnownPath_Inbox Then
        Set oItem = oItem.Move(Folders_KnownPath(glbKnownPath_Deleted))
        If oItem Is Nothing Then Stop: Exit Function
    End If
    
    Set oItem = oItem.Move(oFolder)
    If oItem Is Nothing Then Stop: Exit Function

    '   Delete the file
    '
    glbFSO.DeleteFile FileSpec
    
File_LoadFromFile = True
End Function

' =====================================================================
'   FileSpec Cleanup
' =====================================================================

'   Replace any Invalid Characters in a full FileSpec
'
Public Function File_CleanupFileSpec(ByVal FileSpecRaw As String, Optional ByVal RepChar As String = "_") As String

    Dim Raw As String
    Raw = FileSpecRaw
    
    '   Cut and save the "C:" or "\\" prefix
    '
    Dim Prefix As String
    Prefix = Left(Raw, 2)
    Raw = Mid(Raw, 3)
    
    '   Cleanup the remaining FileSpec
    '
    Dim PathPieces As Variant
    PathPieces = Split(Raw, "\")
    Dim PieceIx As Long
    For PieceIx = 0 To UBound(PathPieces)
        PathPieces(PieceIx) = File_CleanupNameSegment(PathPieces(PieceIx), RepChar)
    Next PieceIx
    
    '   Put the pieces back together with the Prefix
    '
    File_CleanupFileSpec = Prefix & Join(PathPieces, "\")

End Function

'   Replace any Invalid Characters in a Piece of a FileSpec
'
'       NOT for a full File Spec. Only pieces after "C:" and between "/\"s.
'       See: https://learn.microsoft.com/en-us/windows/win32/fileio/naming-a-file
'       2024-10-29 - Added any Control Chars
'
Public Function File_CleanupNameSegment(ByVal Raw As String, Optional ByVal RepChar As String = "_") As String

    Dim Cooked As String
    Cooked = Raw
    
    '   Replace any Control Chars
    '
    Dim AscV As Long
    Dim LoopIx As Long
    For LoopIx = 1 To Len(Cooked)
        AscV = Asc(Mid(Cooked, LoopIx, 1))
        Select Case AscV
            Case 0 To 31, 127, 251 To 255
                Cooked = Replace(Cooked, Chr(AscV), RepChar)
            Case Else
            ' Continue
        End Select
    
    Next LoopIx
    
    '   Replace any Invalids
    '
    Dim Invalids As String
    Invalids = "<>:""/\|?*"

    For LoopIx = 1 To Len(Invalids)
        Cooked = Replace(Cooked, Mid(Invalids, LoopIx, 1), RepChar)
    Next LoopIx

    File_CleanupNameSegment = Cooked

End Function

' =====================================================================
'   Clone
' =====================================================================

'   Create a Clone of an Item in Deleted Items.
'   (So the next Clone.Delete will get rid of it permentley)
'
Public Function File_CloneInDeleted(ByVal Original As Object, ByRef Clone As Object) As Boolean
Const ThisProc = "File_CloneInDeleted"
File_CloneInDeleted = False

    '   Get a Clone in Default Deleted Items
    '
    Dim FileSpec As String
    If Not File_SaveToTemp(Original, FileSpec) Then Stop: Exit Function
    If Not File_LoadFromFile(Clone, FileSpec, Folders_KnownPath(glbKnownPath_Deleted)) Then Stop: Exit Function
    
    '   Delete the Clone (So the next Delete is permanant)
    '   Get it back using it's EntryId
    '
    Dim EntryId As String
    EntryId = Clone.EntryId
    Clone.Delete
    Set Clone = Misc_GetItemFromID(EntryId)
    If Clone Is Nothing Then Stop: Exit Function
    
File_CloneInDeleted = True
End Function

'   Delete an Item Permanently
'
Public Function File_ItemDelete(ByRef Item As Object) As Boolean
File_ItemDelete = False

    If Item Is Nothing Then
        File_ItemDelete = True
        Exit Function
    End If
    
    '   Close with Discard
    '
    Item.Close Outlook.OlInspectorClose.olDiscard
    
    '   Get the Iten's EntryId.
    '   Delete it.
    '   Loop until the EntryId is not found.
    '
    Dim EntryId As String
    Dim LoopCnt As Long
    For LoopCnt = 1 To 3
    
        EntryId = Item.EntryId
        On Error Resume Next
            Item.Delete
            Select Case Err.Number
                Case glbError_None
                Case Else
                    Stop: Exit Function
            End Select
        On Error GoTo 0
        
        Set Item = Nothing
        Set Item = Misc_GetItemFromID(EntryId)
        If Item Is Nothing Then
            File_ItemDelete = True
            Exit Function
        End If
        
    Next LoopCnt
    Stop: Exit Function

End Function
