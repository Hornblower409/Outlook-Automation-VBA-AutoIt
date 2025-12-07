Attribute VB_Name = "Utility_"
Option Explicit
Option Private Module

'   Jump to the Home Inbox Explorer.
'
'   <- Error Msg and Nothing if none found
'
Public Function Utility_HomeInboxExplorer() As Outlook.Explorer
Const ThisProc = "Utility_HomeInboxExplorer"

    '   Return HomeInbox's Explorer
    '
    Set Utility_HomeInboxExplorer = Folders_FolderExplorer(Folders_KnownPath(glbKnownPath_HomeInBox))
    If Utility_HomeInboxExplorer Is Nothing Then
        Msg_Box Proc:=ThisProc, Step:="GetFolderExplorer", Text:="No open Explorer for Folder '" & glbKnownPath_HomeInBox & "'."
        Exit Function
    End If
    
    Utility_HomeInboxExplorer.Activate

End Function

'   Create the Send/Receive Application Folders Group
'
'       So I can disable it before some app adds it and makes it live.
'       From: https://documentation.help/Microsoft-Outlook-Visual-Basic-Reference/olproInAppFolderSyncObject.htm
'
Public Sub Utility_CreateSendReceiveApplicationFoldersGroup()

    Dim olApp As Outlook.Application
    Set olApp = New Outlook.Application
    
    Dim nsp As Outlook.NameSpace
    Dim sycs As Outlook.SyncObjects
    Dim syc As Outlook.SyncObject
    Dim mpfInbox As Outlook.MAPIFolder

    Set nsp = olApp.GetNamespace("MAPI")
    Set sycs = nsp.SyncObjects
    Set syc = sycs.appfolders

End Sub

'   Run External Command using glbAppShell.ShellExecute
'
'       Will process .LNK files in Application.
'       Will handle extensions other than EXE/COM (e.g. A3X) direct or in a LNK.
'       Does NOT expand Environment Variables in Application.
'
Public Sub Utility_ShellExecute( _
    ByVal Application As String, _
    Optional ByVal Parameters As String = "", _
    Optional ByVal WorkingDirectory As String = "", _
    Optional ByVal Verb As String = "", _
    Optional ByVal WindowMode As Integer = 1 _
    )
    
    glbAppShell.ShellExecute Application, Parameters, WorkingDirectory, Verb, WindowMode

End Sub

'   Run an External Command (and Wait) using glbWshShell.Run
'
'       If Wait - Returns the Exit Code of the command (Normally 0 = OK).
'
'           But does not work on LNKs in Command. Returns immediatley. Must be an EXE or COM.
'
'       Will process .LNK files in Command:
'
'           But only of they point to an EXE or COM.
'           Will not work on a LNK to an .A3X.
'
'       Does expand Environment Variables in Command.
'
'
Public Function Utility_ShellRun( _
    ByVal Command As String, _
    Optional ByVal WindowStyle As Integer = 1, _
    Optional ByVal Wait As Boolean = False _
    ) As Integer

    Utility_ShellRun = glbWshShell.Run(Command, WindowStyle, Wait)
    If Not Wait Then Utility_ShellRun = 0

End Function

