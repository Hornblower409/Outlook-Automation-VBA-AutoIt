Attribute VB_Name = "Views_"
Option Explicit
Option Private Module

' =====================================================================
'   All Views
' =====================================================================
'
'   Build a Dict of all Views in All Folders
'
'   Key = Folder Path & vbNewLine & View Name. Item = View Object (as Variant)
'
'   Calling:
'
'       Dim dAllViews As Scripting.Dictionary
'       Set dAllViews = New Scripting.Dictionary
'       If Not Views_AllViews(dAllViews) Then Stop: Exit Function
'
'   Using:
'
'       '   Walk all View items
'       '
'       Dim vView As Variant
'       Dim oView as Outlook.View
'       Dim oFolder As Outlook.Folder
'       For Each vView In dAllViews.Items
'           Set oView = vView
'           Set oFolder = oView.Parent.Parent.Folder
'           ...
'       Next vView
'
'       '   Walk all Keys
'       '
'       Dim Key As Variant
'       Dim KeyParts As Variant
'       For Each Key In dAllViews.Keys
'           KeyParts = Split(Key, vbNewLine)
'           '  KeyParts(0) = Folder Path
'           '  KeyParts(1) = View Name
'       Next Key
'
Public Function Views_AllViews(ByVal dAllViews As Scripting.Dictionary) As Boolean
Views_AllViews = False

    '   Build a Dict of all Folders
    '
    Dim dAllFolders As Scripting.Dictionary
    Set dAllFolders = New Scripting.Dictionary
    If Not Folders_AllFolders(dAllFolders) Then Stop: Exit Function
    
    '   Build a Dict of all Views in all Folders
    '
    Dim oFolder As Outlook.Folder
    Dim vFolder As Variant
    Dim oView As Outlook.View
    For Each vFolder In dAllFolders.Items
        Set oFolder = vFolder
        For Each oView In oFolder.Views
            dAllViews.Add oFolder.FolderPath & vbNewLine & oView.Name, oView
        Next oView
    Next vFolder

Views_AllViews = True
End Function

' =====================================================================
'   Views Save All
' =====================================================================
'
Public Sub Views_SaveAll()
Const ThisProc = "Views_SaveAll"

    '   Get a File Folder Path to Save in and the Overwrite switch
    '
    Dim SaveFolderPath As String
    Dim Overwrite As Boolean
    
    If Not File_FolderSelectOverwrite( _
        SaveFolderPath, _
        Overwrite, _
        Caption:="Select or Create a Save All Views folder", _
        FolderStartPath:=glbViewsStartFolder, _
        Options:=&H10 _
    ) Then Exit Sub

    '   Get a Dict of all Views in all Folders
    '
    Dim dAllViews As Scripting.Dictionary
    Set dAllViews = New Scripting.Dictionary
    If Not Views_AllViews(dAllViews) Then Stop: Exit Sub
    
    '   Get a count of how many Folders in the Dict
    '
    Dim Key As Variant
    Dim KeyParts As Variant
    Dim LastFolderPath As String: LastFolderPath = ""
    Dim FolderCount As Long: FolderCount = 0
    For Each Key In dAllViews.Keys
        KeyParts = Split(Key, vbNewLine)
        If LastFolderPath <> KeyParts(0) Then
            FolderCount = FolderCount + 1
            LastFolderPath = KeyParts(0)
        End If
    Next Key
    
    '   Write each View as a File in the Save File Folder.
    '   Creating file subfolders for each Outlook folder.
    '
    Dim vView As Variant
    Dim oView As Outlook.View
    Dim Skipped As Boolean
    Dim SkippedCount As Long: SkippedCount = 0
    Dim SavedCount As Long: SavedCount = 0
    For Each vView In dAllViews.Items
        Set oView = vView
        If Not Views_SaveAsXML(oView, SaveFolderPath, Overwrite, Skipped) Then Stop: Exit Sub
        If Skipped Then SkippedCount = SkippedCount + 1 Else SavedCount = SavedCount + 1
    Next vView
    
    '   Tell them what I've done
    '
    Msg_Box _
        Proc:=ThisProc, _
        Icon:=vbInformation, _
        Text:="Save Complete." & glbBlankLine & _
        FolderCount & " Folders processed. " & dAllViews.count & " Views processed." & vbNewLine & _
        SavedCount & " Views saved. " & SkippedCount & " Views skipped."

End Sub

'   Write a View's XML as a file to a folder
'
Private Function Views_SaveAsXML( _
    ByVal oView As Outlook.View, _
    ByVal SaveFolderPath As String, _
    ByVal Overwrite As Boolean, _
    ByRef Skipped As Boolean _
    ) As Boolean
Const ThisProc = "Views_SaveAsXML"
Views_SaveAsXML = False

    '   Build a FileSpec from the Outlook Folder Path \ View Name
    '
    '       Outlook FolderPath is something like "\\{Store}\{Folder}...\{Folder}"
    '       View Name can be anything!
    '
    
    '   Get the View Folder Path.
    '
    '       (Parent.Parent is because the immediate parent of a View is the Folder.Views collection).
    '
    Dim FolderPath As String
    FolderPath = oView.Parent.Parent.FolderPath
    
    '   Strip the leading "\\" from the FolderPath.
    '   Replace any invalid File name chars in each piece of the path.
    '
    '   (He will have URL encoded any "\/%" in the Folder Names).
    '
    FolderPath = Mid(FolderPath, 3)
    Dim PathPieces As Variant
    PathPieces = Split(FolderPath, "\")
    Dim PieceIx As Long
    For PieceIx = 0 To UBound(PathPieces)
        PathPieces(PieceIx) = File_CleanupNameSegment(PathPieces(PieceIx))
    Next PieceIx
    FolderPath = Join(PathPieces, "\")
    
    '   Append the SaveFolderPath
    '
    '   (No "\" needed as SaveFolderPath ends with a "\")
    '
    FolderPath = SaveFolderPath & FolderPath
    
    '   Build the tail end of the File Spec (the View Name)
    '   Replace any invalid File name chars in the View Name
    '   Build a full FileSpec for the File
    '
    Dim ViewName As String
    ViewName = File_CleanupNameSegment(oView.Name) & ".xml"
    Dim FileSpec As String
    FileSpec = FolderPath & "\" & ViewName

    '   If the file exist and no overwrite - done
    '
    Skipped = False
    If glbFSO.FileExists(FileSpec) And Not Overwrite Then
        Skipped = True
        Views_SaveAsXML = True
        Exit Function
    End If

    '   Create the Folder Path (all the way down) if it doesn't already exist
    '
    If Not File_MkDir(FolderPath) Then Stop: Exit Function
    
    '   Build my comment line: <!-- Date Time GUID  OutlookFolderPath ViewName -->
    '   with OutlookFolderPath and ViewName Hex Encoded.
    '
    Dim DateTime As String
    DateTime = Misc_NowString()
    Dim OutlookFolderPathHEX As String
    If Not Misc_PlainToHex(oView.Parent.Parent.FolderPath, OutlookFolderPathHEX) Then Stop: Exit Function
    Dim ViewNameHEX As String
    If Not Misc_PlainToHex(oView.Name, ViewNameHEX) Then Stop: Exit Function
    Dim Comment As String
    Comment = "<!-- " & DateTime & " " & glbViewsSaveGUID & " " & OutlookFolderPathHEX & " " & ViewNameHEX & " -->"
    
    '   Get the View XML and insert my Comment as the 2nd line.
    '
    Dim ViewXML As String
    ViewXML = oView.xml
    Dim Index As Long
    Index = InStr(ViewXML, vbNewLine)
    ViewXML = Mid(ViewXML, 1, Index + 1) & Comment & vbNewLine & Mid(ViewXML, Index + 2)    ' !!  vbNewLine is TWO chars long = vbCr & vbLf  !!
    
    '   Write it
    '
    If Not File_WriteText(FileSpec, ViewXML) Then Exit Function

Views_SaveAsXML = True
End Function
