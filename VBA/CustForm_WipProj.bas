Attribute VB_Name = "CustForm_WipProj"
Option Explicit
Option Private Module

' =====================================================================
'   Misc
' =====================================================================

    '   ID Seq Value
    '
    Private Const SeqRegKey         As String = glbRegOutlookBaseKey & "Projects\WIPProject\Seq"
    Private Const SeqInit           As Long = 1000
    Private Const SeqMax            As Long = 9999
    
    '   WIP Dir
    '
    Private Const WipDirBasePath    As String = "C:\DATA\PROJECTS\"
    
    '   WIP Master Cat
    '
    Private Const MasterCatColor    As Long = Outlook.olCategoryColorYellow
    
' =====================================================================
'   WIP Status That Can Be Archived Table
' =====================================================================

    Private Const WipStatusClosed   As String = "Closed"
    Private Const WipStatusCanceled As String = "Canceled"

    Private Enum WipStatusArchiveTableCol
        wsaStatus                   ' Status Value
    End Enum

    Private Const WipStatusArchiveTableConst As String = _
 _
    "     Status                " & vbLf & _
    "'    ----------------------" & vbLf & _
    " " & WipStatusClosed & "   " & vbLf & _
    " " & WipStatusCanceled & " " & vbLf & _
    "'"
    
' =====================================================================
'   Master Cat Update Action Table
' =====================================================================

    Private Enum MasterCatAction
        NoOp
        Add
        Rename
        Collision
    End Enum
    
    Private Enum MasterCatActionTableCol
        mccKey
        mccAction
    End Enum
    
    Private Const MasterCatActionTableConst As String = _
 _
    "'                                        |  NewCat =  |  NewCat   |  OldCat    " & vbLf & _
    "  Key  |  Action                         |  OldCat?   |  Exist?   |  Exist?    " & vbLf & _
    "' ---- |  ------------------------------ |  --------- |  -------- |  --------- " & vbLf & _
    "  000  |" & MasterCatAction.Add & "      |  N         |  N        |  N         " & vbLf & _
    "  001  |" & MasterCatAction.Rename & "   |  N         |  N        |  Y         " & vbLf & _
    "  010  |" & MasterCatAction.NoOp & "     |  N         |  Y        |  N         " & vbLf & _
    "  011  |" & MasterCatAction.Collision & "|  N         |  Y        |  Y         " & vbLf & _
    "  100  |" & MasterCatAction.Add & "      |  Y         |  N        |  N         " & vbLf & _
    "  101  |" & MasterCatAction.NoOp & "     |  Y         |  N        |  Y         " & vbLf & _
    "  110  |" & MasterCatAction.NoOp & "     |  Y         |  Y        |  N         " & vbLf & _
    "  111  |" & MasterCatAction.NoOp & "     |  Y         |  Y        |  Y         " & vbLf & _
    "'"

' =====================================================================
'   WipDir Update Action Table
' =====================================================================

    Private Enum WipDirAction
        NoOp
        Add
        Rename
        Collision
    End Enum
    
    Private Enum WipDirActionTableCol
        wdcKey
        wdcAction
    End Enum
    
    Private Const wipDirActionTableConst As String = _
 _
    "'                                     |  NewDir =  |  NewDir   |  OldDir    " & vbLf & _
    "  Key  |  Action                      |  OldDir?   |  Exist?   |  Exist?    " & vbLf & _
    "' ---- |  ----------------------------|  --------- |  -------- |  --------- " & vbLf & _
    "  000  |" & WipDirAction.Add & "      |  N         |  N        |  N         " & vbLf & _
    "  001  |" & WipDirAction.Rename & "   |  N         |  N        |  Y         " & vbLf & _
    "  010  |" & WipDirAction.NoOp & "     |  N         |  Y        |  N         " & vbLf & _
    "  011  |" & WipDirAction.Collision & "|  N         |  Y        |  Y         " & vbLf & _
    "  100  |" & WipDirAction.Add & "      |  Y         |  N        |  N         " & vbLf & _
    "  101  |" & WipDirAction.NoOp & "     |  Y         |  N        |  Y         " & vbLf & _
    "  110  |" & WipDirAction.NoOp & "     |  Y         |  Y        |  N         " & vbLf & _
    "  111  |" & WipDirAction.NoOp & "     |  Y         |  Y        |  Y         " & vbLf & _
    "'"

' =====================================================================
'   CmdButton Table
' =====================================================================

    Private Const cbOpenWipDir      As String = "WIPOpenWIPDirButton"
    Private Const cbArchive         As String = "WIPArchiveButton"

    Private Enum CmdButtonTableCol
        cbcControlName              '   Control Internal Name.
    End Enum

    Private Const CmdButtonTableConst As String = _
 _
    "'    Control           " & vbLf & _
    "'    Internal Name     " & vbLf & _
    "'                      " & vbLf & _
    "     CtrlName          " & vbLf & _
    "'    ------------------" & vbLf & _
    " " & cbOpenWipDir & "  " & vbLf & _
    " " & cbArchive & "     " & vbLf & _
    "'"
    
' =====================================================================
'   User Prop Names
' =====================================================================
    
    Private Const pnDescription     As String = "WIPDescription"
    Private Const pnDocType         As String = "WIPDocType"
    Private Const pnDueDate         As String = "WIPDueDate"
    Private Const pnID              As String = "WIPID"
    Private Const pnMasterCat       As String = "WIPMasterCat"
    Private Const pnMasterCatSaved  As String = "WIPMasterCatSaved"
    Private Const pnPriority        As String = "WIPPriority"
    Private Const pnSeq             As String = "WIPSeq"
    Private Const pnStatus          As String = "WIPStatus"
    Private Const pnStatusUpdate    As String = "WIPStatusUpdate"
    Private Const pnTitle           As String = "WIPTitle"
    Private Const pnType            As String = "WIPType"
    Private Const pnWipDir          As String = "WIPDir"
    Private Const pnWipDirSaved     As String = "WIPDirSaved"
    
    '   Shadows
    '
    '   WIPFollowUpText
    ''   Computed Shadow of .FlagRequest [Follow Up Flag]
    ''   IIf("X" & [Follow Up Flag] & "X"="XX","",ChrW(9654) & " " & [Follow Up Flag] )
    '
    '   WIPReminderTime
    ''   Computed Shadow of .ReminderSet [Reminder], .ReminderTime [Reminder Time]
    ''   IIF([Reminder], Format([Reminder Time], "ddd" ) & " " & Format([Reminder Time], "yyyy-mm-dd") & " " & Format([Reminder Time] , "hh:nn"), "")
    
' =====================================================================
'   Field Table
' =====================================================================
   
    Private Enum PropType
        ptStd
        ptUser
    End Enum

    Private Enum FieldTableCol
        ftcFieldName            '   Prop/Ctrl Name.
        ftcPropType             '   Prop Type from Enum PropType.
        ftcDisplayName          '   Field Display Name.
        ftcRequired             '   Is Required? Can not be "" for Text/Combo, "" or 0 for Number.
        ftcMin                  '   MinLen for Text/Combo. MinValue for Number.
        ftcMax                  '   MaxLen for Text/Combo. MaxValue for Number.
    End Enum

    Private Const FieldTableConst As String = _
 _
    "' Prop/Ctrl         | Prop         | Field         | Value     |     |     " & vbLf & _
    "' Name              | Type         | Display Name  | Required? |     |     " & vbLf & _
    "'                   |              |               |           |     |     " & vbLf & _
    "  FieldName         | PropType     | DisplayName   | Required  | Min | Max " & vbLf & _
    "' ----------------- | ------------ | ------------- | --------- | --- | --- " & vbLf & _
       pnTitle & "       |" & ptUser & "| Title         | True      |   1 |  48 " & vbLf & _
       pnType & "        |" & ptUser & "| Type          | True      |     |     " & vbLf & _
       pnID & "          |" & ptUser & "| ID            | True      |     |     " & vbLf & _
       pnDescription & " |" & ptUser & "| Desc          | False     |     |     " & vbLf & _
       pnStatus & "      |" & ptUser & "| Status        | True      |     |     " & vbLf & _
       pnStatusUpdate & "|" & ptUser & "| Status Update | False     |   1 |  64 " & vbLf & _
       pnPriority & "    |" & ptUser & "| Priority      | True      |     |     " & vbLf & _
       pnDueDate & "     |" & ptUser & "| Due           | False     |     |     " & vbLf & _
    "'"

' =====================================================================
'   Custom Action
' =====================================================================

Public Function CustForm_WipProjCustAction(ByVal oForm As Outlook.PostItem, ByVal Action As String) As Boolean
CustForm_WipProjCustAction = False


CustForm_WipProjCustAction = True
End Function

' =====================================================================
'   Prop Change
' =====================================================================

Public Function CustForm_WipProjPropChange(ByVal oForm As Outlook.PostItem, ByVal IsStandardProp As Boolean, ByVal PropName As String) As Boolean
CustForm_WipProjPropChange = False

    Select Case True
        Case IsStandardProp
            If Not CustForm_WipProjPropChangeStd(oForm, PropName) Then Stop: Exit Function
        Case Else
            If Not CustForm_WipProjPropChangeUser(oForm, PropName) Then Stop: Exit Function
    End Select

CustForm_WipProjPropChange = True
End Function

Private Function CustForm_WipProjPropChangeStd(ByVal oForm As Outlook.PostItem, ByVal PropName As String) As Boolean
CustForm_WipProjPropChangeStd = False

    Select Case PropName
        Case Else
            '   Continue
    End Select

CustForm_WipProjPropChangeStd = True
End Function

Private Function CustForm_WipProjPropChangeUser(ByVal oForm As Outlook.PostItem, ByVal PropName As String) As Boolean
CustForm_WipProjPropChangeUser = False

    Select Case PropName
        Case pnType
            If Not CustForm_WipProjRecalc(oForm) Then Stop: Exit Function
        Case pnTitle
            If Not CustForm_WipProjRecalc(oForm) Then Stop: Exit Function
        Case Else
            '   Continue
    End Select
    
CustForm_WipProjPropChangeUser = True
End Function

' =====================================================================
'   Open
' =====================================================================

Public Function CustForm_WipProjOpen(ByVal oForm As Outlook.PostItem) As Boolean
CustForm_WipProjOpen = False

    '   Hook my Command Buttons
    '
    If Not CustForm_WipProjHookCmdButtons(oForm) Then Stop: Exit Function
    
    '   If a new Item
    '
    If (Len(oForm.EntryId) = 0) Then
    
        '   Generate a New Seq
        '
        If Not CustForm_WipProjSeq(oForm) Then Stop: Exit Function
        
        '   Set ExpiryTime = My special marker value
        '
        '   For X1 Search. WIP Project items must have the Expires field set to a Max Value
        '   so they will sort out at the top. It was the only field (other than Flag) that I
        '   could find that is in X1 and I can play with it in VBA code.
        '
        If oForm.ExpiryTime = glbDateNone Then
            oForm.ExpiryTime = glbDateExpiresFlag
        End If
        
    End If
    
    '   If an old V3 that has no Seq field - Extract Seq from ID
    '
    If UserProp_Get(oForm, pnSeq) = "" Then
        UserProp_Set oForm, pnSeq, Mid(UserProp_Get(oForm, pnID), 4)
        oForm.Save
    End If

CustForm_WipProjOpen = True
End Function

' =====================================================================
'   Command Buttons
' =====================================================================

Public Function CustForm_WipProjHookCmdButtons(ByVal oForm As Outlook.PostItem) As Boolean
CustForm_WipProjHookCmdButtons = False

    Dim InspectorRec As Inspector_InspectorRec
    If Not Inspector_RecFromItem(oForm, InspectorRec) Then Stop: Exit Function
    
    With InspectorRec
    
        '   Load the CmdButton Table
        '
        Dim CmdButtonList() As String
        CmdButtonList = Tbl_TableConstList(CmdButtonTableConst)

        '   Get each Control by Name and Hook it
        '
        Dim RowIndex As Long
        For RowIndex = LBound(CmdButtonList) To UBound(CmdButtonList)
    
            Dim oCmdButton As MSForms.Control
            If Not CustForm_ControlByName(oForm, CmdButtonList(RowIndex), oCmdButton) Then Stop: Exit Function
            .oInspShadow.CmdButtonHook oCmdButton
        
        Next RowIndex
        
    End With

CustForm_WipProjHookCmdButtons = True
End Function

Public Function CustForm_WipProjClickCmdButton(ByVal oForm As Outlook.PostItem, ByVal oCmdButton As MSForms.CommandButton) As Boolean
Const ThisProc = "CustForm_WipProjClickCmdButton"
CustForm_WipProjClickCmdButton = False

    Select Case oCmdButton.Name
        Case cbOpenWipDir
            GoSub CmdButton_OpenWipDir
        Case cbArchive
            If Not CustForm_WipProjArchive(oForm) Then Stop: Exit Function
        Case Else
            Stop: Exit Function
    End Select

CustForm_WipProjClickCmdButton = True
Exit Function

'   CmdButton - Open Wip Dir
'
CmdButton_OpenWipDir:

    '   Form must be Saved
    '
    If Not oForm.Saved Then
        Msg_Box Proc:=ThisProc, Step:="OpenWipDir - Form Saved?", _
            Text:="Form must be saved (or at least has found Jesus) before you can open it's WIP Dir."
        CustForm_WipProjClickCmdButton = True
        Exit Function
    End If
    
    '   Wip Dir must be defined
    '
    Dim WipDir As String
    WipDir = UserProp_Get(oForm, pnWipDirSaved)
    If WipDir = "" Then Stop: Exit Function

    '   If Wip Dir doesn't exist - Create it.
    '
    If Not glbFSO.FolderExists(WipDirBasePath & WipDir) Then
        If Not File_MkDir(WipDirBasePath & WipDir) Then Stop: Exit Function
    End If

    '   Call HotRod Explorer_OpenPath to open the WIP Dir
    '
    Utility_ShellExecute _
        Application:=glbHotRodLnks & glbHotRodLnks_ExplorerOpenPath, _
        Parameters:="/ActONE " & glbQuote & WipDirBasePath & WipDir & glbQuote
        
    '   Put the quoted path on the Clipboard
    '
    Misc_ClipSet glbQuote & WipDirBasePath & WipDir & glbQuote

Return

End Function

' =====================================================================
'   Archive
' =====================================================================

Private Function CustForm_WipProjArchive(ByVal oForm As Outlook.PostItem) As Boolean
Const ThisProc = "CustForm_WipProjArchive"
CustForm_WipProjArchive = False

    GoSub Archive_Validate
    GoSub Archive_Filter
    GoSub Archive_Execute

Archive_Exit: CustForm_WipProjArchive = True
Exit Function

'   Archive Validate
'
Archive_Validate:

    '   Must be Saved
    '
    If Not oForm.Saved Then
        Msg_Box Proc:=ThisProc, Step:="Archive_Validate - Form Saved?", _
            Text:="Form must be saved (or at least has found Jesus) before you can Archive it."
        GoTo Archive_Exit
    End If
    
    '   Must not already be Archived
    '
    If oForm.Parent.FolderPath = glbKnownPath_ProjectsArchive Then
        Msg_Box Proc:=ThisProc, Step:="Archive_Validate - Already Archived?", _
            Text:="Form is already in the Archive Folder."
        GoTo Archive_Exit
    End If
    
    '   Must have a Status that can be Archived
    '
    Dim WipStatus As String
    WipStatus = UserProp_Get(oForm, pnStatus)
    If Not Tbl_TableConstExist(WipStatusArchiveTableConst, WipStatus) Then
        Msg_Box Proc:=ThisProc, Step:="Archive_Validate - Status can be Archived?", _
            Text:="WIP Status: '" & WipStatus & "'." & glbBlankLine & _
                  "Form Status must be '" & Join(Tbl_TableConstList(WipStatusArchiveTableConst), ", ") & "' to Archive."
        GoTo Archive_Exit
    End If
    
    '   WIP Master Cat must exist in the Master Cat List
    '
    Dim MasterCat As String
    MasterCat = UserProp_Get(oForm, pnMasterCatSaved)
    
    If Not Categories_MasterCatsExist(MasterCat) Then
        Msg_Box Proc:=ThisProc, Step:="Archive_Validate - WIP Master Cat Exist?", _
            Text:="Form's WIP Master Cat: '" & MasterCat & "'." & glbBlankLine & _
                  "Does not exist in the Master Cat List. (Edit and Save?)"
        GoTo Archive_Exit
    End If
    
    '   Form must be assigned it's own WIP Master Cat
    '
    If Not Categories_FindCat(oForm, MasterCat) Then
        Msg_Box Proc:=ThisProc, Step:="Archive_Validate - WIP Master Cat Assigned?", _
            Text:="Form's WIP Master Cat: '" & MasterCat & "'." & glbBlankLine & _
                  "Is not assigned to the Form."
        GoTo Archive_Exit
    End If
    
    '   Get a CatsList without any WIP Cats or Specials
    '
    Dim CatsList As String
    CatsList = oForm.Categories
    CatsList = Categories_RemoveCatList(CatsList, MasterCat)
    CatsList = Categories_RemoveSpecialCatsList(CatsList)
    
    '   CatsList must have at least one "Normal" (Not Special. Not WIP) Cat
    '
    If Len(Trim(CatsList)) = 0 Then
        Msg_Box Proc:=ThisProc, Step:="Archive_Validate - Normal Cat?", _
            Text:="Form does not have a 'Normal' (not Special, not WIP) Cat."
        GoTo Archive_Exit
    End If
    
    '   Can not have any Active Foreign WIP Master Cats
    '
    Dim ForeignMaster As String
    Dim oItem As Object
    Set oItem = oForm
    GoSub Archive_ForeignMaster
    If Not ForeignMaster = "" Then
        Msg_Box Proc:=ThisProc, Step:="Archive_Validate - Active Foreign WIP Master Cats?", Text:= _
            "Form WIP Master Cat: '" & MasterCat & "'." & glbBlankLine & _
            "Foreign WIP Master Cat: '" & ForeignMaster & "'." & glbBlankLine & _
            "Form can not have any Active Foreign WIP Master Cat assigned."
        GoTo Archive_Exit
    End If
    
    '   Stash the Form's Normal Cats List
    '
    Dim NormalCatsList As String
    NormalCatsList = CatsList
    
    '   Stash the Form's WIP Dir
    '
    Dim WipDirSaved As String
    WipDirSaved = UserProp_Get(oForm, pnWipDirSaved)
    
Return

'   Archive Filter
'
Archive_Filter:

    '   Get a Collection of this Form and all Project Items with this WIP Master Cat
    '
    Dim Filter As String
    Dim Related As VBA.Collection
    
    If Not Categories_CatFilter(MasterCat, Filter) Then Stop: Exit Function
    Filter = "@SQL=" & Filter
    If Not Collection_FromRestrict(Filter, Folders_KnownPath(glbKnownPath_Projects), Related) Then Stop: Exit Function

    '   Walk the Related Collection Backwards
    '
    Dim SkippedForeignMasterCat As Long
    Dim RelatedIx As Long
    For RelatedIx = Related.count To 1 Step -1: Do
            
        Set oItem = Related.Item(RelatedIx)
        
        '   Must be Saved
        '
        If Not oItem.Saved Then
            Msg_Box Proc:=ThisProc, Step:="Archive_Filter - Item Saved?", Text:= _
                "Item: '" & oItem.Subject & "'." & vbNewLine & _
                "MsgClass: '" & oItem.MessageClass & "'." & vbNewLine & _
                "EntryId: '" & oItem.EntryId & "'." & glbBlankLine & _
                "Is not Saved."
            GoTo Archive_Exit
        End If
        
        '   Can not have a Follow Up Flag
        '
        Dim FlagRequest As String
        If Not FollowUp_GetFlagRequest(oItem, FlagRequest) Then Stop: Exit Function
        
        If Not FlagRequest = "" Then
            Msg_Box Proc:=ThisProc, Step:="Archive_Filter - Flag Request?", _
                Text:="Item: '" & oItem.Subject & "'." & glbBlankLine & _
                "Has a Flag Request (Follow Up Text) value."
            GoTo Archive_Exit
        End If
        
        '   Can not have a Reminder Set
        '
        If oItem.ReminderSet Then
            Msg_Box Proc:=ThisProc, Step:="Archive_Filter - Reminder Set?", _
                Text:="Item: '" & oItem.Subject & "'." & glbBlankLine & _
                "Has a Reminder set."
            GoTo Archive_Exit
        End If
        
        '   Can not have a Follow Up Cat
        '
        If Categories_FindCat(oItem, glbCatFollowUp) Then
            Msg_Box Proc:=ThisProc, Step:="Archive_Filter - Follow Up Cat?", _
                Text:="Item: '" & oItem.Subject & "'." & glbBlankLine & _
                "Has a Follow Up Cat."
            GoTo Archive_Exit
        End If
        
        '   Can not be another WIP Project
        '
        If oItem.MessageClass = glbCustForm_WipProject Then
            If Not Session.CompareEntryIDs(oItem.EntryId, oForm.EntryId) Then
                Msg_Box Proc:=ThisProc, Step:="Archive_Filter - Another WIP Project?", Text:= _
                    "Item EntryId: '" & oItem.EntryId & "'." & vbNewLine & _
                    "Form EntryId: '" & oForm.EntryId & "'." & glbBlankLine & _
                    "WIP Master Cat: '" & MasterCat & "'." & glbBlankLine & _
                    "Duplicate WIP Projects with the same WIP Master Cat."
                GoTo Archive_Exit
            End If
        End If
        
        '   If oItem has an Active Foreign WIP Master Cats - Skip
        '
        GoSub Archive_ForeignMaster
        If Not ForeignMaster = "" Then
            Related.Remove RelatedIx
            SkippedForeignMasterCat = SkippedForeignMasterCat + 1
            Exit Do     ' Next RelatedIx
        End If
        
    Loop While False: Next RelatedIx

    '   If nothing to Archive - Bail
    '
    If Related.count = 0 Then
        Msg_Box Proc:=ThisProc, Step:="Archive_Filter - Nothing to Archive?", Text:= _
            "WIP Master Cat: '" & MasterCat & "'." & glbBlankLine & _
            "Form and all of it's related Items can not be Archived."
        GoTo Archive_Exit
    End If
    
Return

'   Count the number of Active Foreign WIP Master Cats in oItem
'
Archive_ForeignMaster:

    ForeignMaster = ""
    
    '   Walk the Item Cats
    '
    Dim Cats() As String
    Cats = Split(oItem.Categories, glbCatSep)
    
    Dim Cat As String
    Dim CatIx As Long
    For CatIx = LBound(Cats) To UBound(Cats): Do
    
        Cat = Trim(Cats(CatIx))

        '   If Cat is the Form's WIP Master Cat - Ignore
        '
        If Cat = MasterCat Then Exit Do ' Next CatIx
        
        '   If Cat is a Active Foreign WIP Master Cat - Return it
        '
        If InStr(1, Cat, WIPCatPrefix) = 1 Then
            If Categories_MasterCatsExist(Cat) Then
                ForeignMaster = Cat
                Return
            End If
        End If
    
    Loop While False: Next CatIx

Return

'   Archive Execute
'
Archive_Execute:
    
    '   Archive Show and Tell and Get Confirmation
    '
    Select Case Msg_Box( _
        Proc:=ThisProc, Step:="Archive_Execute - Confirm?", _
        Icon:=vbQuestion, Buttons:=vbYesNoCancel, Default:=vbDefaultButton2, _
        Text:= _
            "WIP Master Cat: '" & MasterCat & "'." & glbBlankLine & _
            "Skipped - Foreign Active WIP Master Cats: " & SkippedForeignMasterCat & glbBlankLine & _
            "Form and " & Related.count - 1 & " related items. " & vbNewLine & _
            "Archive to: '" & glbKnownPath_ProjectsArchive & "'." & glbBlankLine & _
            "Continue?")
        Case vbYes
            ' Continue
        Case Else
            GoTo Archive_Exit
    End Select
    
    '   Walk the Related Items Collection
    '
    For RelatedIx = 1 To Related.count
            
        '   Get the Item and Close any Open Inspectors
        '
        Set oItem = Related.Item(RelatedIx)
        oItem.Close OlInspectorClose.olSave
    
        '   Add the Normal Cats from the Form
        '   Remove any Special Cats (Leave any Priority Cat)
        '   Save
        '
        Categories_AddCats oItem, NormalCatsList
        Categories_RemoveSpecialCats oItem, PriorityCats:=False
        oItem.Save
        
        '   Move it to the Archive
        '
        Dim oArchived As Object
        Set oArchived = oItem.Move(Folders_KnownPath(glbKnownPath_ProjectsArchive))
        If oArchived Is Nothing Then Stop: Exit Function
    
    Next RelatedIx
    
    '   Remove the Form's WIP Cat from Master Cats
    '
    If Not Categories_MasterCatsRemove(MasterCat) Then Stop: Exit Function
    
    '   If the WIP Dir exist and is empty - Delete it
    '
    If glbFSO.FolderExists(WipDirBasePath & WipDirSaved) Then
        Dim oWipDirSaved As Scripting.Folder
        Set oWipDirSaved = glbFSO.GetFolder(WipDirBasePath & WipDirSaved)
        If oWipDirSaved.Files.count = 0 And oWipDirSaved.SubFolders.count = 0 Then
            oWipDirSaved.Delete
        End If
    End If
    
Return

End Function

' =====================================================================
'   Write
' =====================================================================

Public Function CustForm_WipProjWrite(ByVal oForm As Outlook.PostItem) As Boolean
Const ThisProc = "CustForm_WipProjWrite"
CustForm_WipProjWrite = False

    ' 2025-06-01 - If already Saved - Bail
    '
    '   Mail_ExportAsPDF triggers a Write when he does a SaveAs.
    '   Which causes this Proc to redo stuff and make the Original Dirty.
    '   But since it's a Save AS, the Original is never Saved.
    '
    If oForm.Saved Then
        CustForm_WipProjWrite = True
        Exit Function
    End If
    
    '   Load the Field Table
    '
    Dim FieldTableList() As String
    FieldTableList = Tbl_TableConstList(FieldTableConst)

    '   Walk the Field Table and Validate
    '
    Dim RowIndex As Long
    For RowIndex = LBound(FieldTableList) To UBound(FieldTableList)
        If Not CustForm_WipProjValidate(oForm, FieldTableList(RowIndex)) Then Exit Function
    Next RowIndex
    
    '   Title must usable in a File Name
    '
    GoSub Write_TitleCheck
    
    '   Assign Priority and Status Cats
    '
    GoSub Write_AssignCats
    
    '   Master Cat and Wip Dir Updates
    '
    If Not CustForm_WipProjUpdate(oForm) Then Exit Function
    
    '   Update Saved Values
    '
    UserProp_Set oForm, pnMasterCatSaved, UserProp_Get(oForm, pnMasterCat)
    UserProp_Set oForm, pnWipDirSaved, UserProp_Get(oForm, pnWipDir)

    '   Update my Master Cat
    '
    If Not Categories_MasterCatsAssign(oForm, UserProp_Get(oForm, pnMasterCat)) Then Stop: Exit Function
    
CustForm_WipProjWrite = True
Exit Function

'   Check Title can be used in a File Name
'
Write_TitleCheck:

    Dim FileSpecClean As String
    FileSpecClean = File_CleanupFileSpec(UserProp_Get(oForm, pnWipDir))
    
    If (UserProp_Get(oForm, pnWipDir) = FileSpecClean) _
    And (InStr(1, UserProp_Get(oForm, pnWipDir), ",") = 0) _
    Then Return

    Msg_Box Proc:=ThisProc, Step:="Check Title", _
        Text:="Title can not contain any characters that are not valid in a Windows File Name or comma."
    Exit Function

Return

'   Assign Cats based on Priority, Status, Follow Up
'
Write_AssignCats:

    '   Remove any existing Priority Cats
    '
    Categories_RemovePriorityCats oForm
    
    '   Priority Cat <- Form Priority Property with a Prefix
    '
    Dim PriorityCat As String
    PriorityCat = glbCatPrefixPriority & " " & UserProp_Get(oForm, pnPriority)
    
    '   If Priority Cat is a Priority Master Cat - Assign it.
    '
    Select Case PriorityCat
        Case glbCatPriorityHigh, glbCatPriorityMedium, glbCatPriorityLow
            Categories_AddCat oForm, PriorityCat
        Case Else
            Stop: Exit Function
    End Select

    '   For selected Form Status Property values - Add a Wait Priority Master Cat
    '
    Select Case UserProp_Get(oForm, pnStatus)
        Case "Stalled", "Waiting"
            Categories_AddCat oForm, glbCatPriorityWait
        Case Else
            ' Continue
    End Select
    
    '   Update Follow Up Cat
    '
    Categories_RemoveCat oForm, glbCatFollowUp
    Dim FlagRequest As String
    If Not FollowUp_GetFlagRequest(oForm, FlagRequest) Then Stop: Exit Function
    If Not FlagRequest = "" Then Categories_AddCat oForm, glbCatFollowUp

Return

End Function

' ---------------------------------------------------------------------
'   Write - Update
' ---------------------------------------------------------------------
'
Private Function CustForm_WipProjUpdate(ByVal oForm As Outlook.PostItem) As Boolean
Const ThisProc = "CustForm_WipProjUpdate"
CustForm_WipProjUpdate = False

    '   Get Update Actions
    '
    Dim NewCat As String: Dim OldCat As String: Dim MasterCatUpdateAction As Long
    Dim NewDir As String: Dim OldDir As String: Dim WipDirUpdateAction As Long
    GoSub Update_GetUpdateActions

    '   Collision
    '
    If MasterCatUpdateAction = MasterCatAction.Collision Then
        Msg_Box Proc:=ThisProc, Step:="MasterCat Update Collision", _
            Text:="Master Cat:" & glbBlankLine & _
                  NewCat & glbBlankLine & _
                  "Already exist."
        Exit Function
    End If
    
    If WipDirUpdateAction = WipDirAction.Collision Then
        Msg_Box Proc:=ThisProc, Step:="Wip Dir Update Collision", _
            Text:="WIP Dir:" & glbBlankLine & _
                  WipDirBasePath & NewDir & glbBlankLine & _
                  "Already exist."
        Exit Function
    End If
    
    '   Add
    '
    If MasterCatUpdateAction = MasterCatAction.Add Then
        If Not Categories_MasterCatsAdd(NewCat, MasterCatColor) Then Stop: Exit Function
    End If
    
    '   2025-05-24 - Don't create until opened
    '
    ' If WipDirUpdateAction = WipDirAction.Add Then
    '    If Not File_MkDir(WipDirBasePath & NewDir) Then Stop: Exit Function
    ' End If
    
    '   Rename
    '
    If MasterCatUpdateAction = MasterCatAction.Rename Then
        Categories_RemoveCat oForm, OldCat
        If Not Categories_MasterCatsRename(OldCat, NewCat) Then Stop: Exit Function
    End If
    If WipDirUpdateAction = WipDirAction.Rename Then
        On Error Resume Next
            glbFSO.MoveFolder WipDirBasePath & OldDir, WipDirBasePath & NewDir
        On Error GoTo 0
        If Not Err.Number = glbError_None Then Stop: Exit Function
    End If
    
CustForm_WipProjUpdate = True
Exit Function

'   Get Update Actions
'
Update_GetUpdateActions:

    '   Master Cat
    '
    NewCat = UserProp_Get(oForm, pnMasterCat)
    OldCat = UserProp_Get(oForm, pnMasterCatSaved)

    Dim MasterCatActionKey As String
    MasterCatActionKey = _
        IIf(NewCat = OldCat, "1", "0") & _
        IIf(Categories_MasterCatsExist(NewCat), "1", "0") & _
        IIf(Categories_MasterCatsExist(OldCat), "1", "0")
        
    Dim MasterCatActionStr As String
    If Not Tbl_TableConstFind(MasterCatActionTableConst, MasterCatActionKey, mccAction, MasterCatActionStr) Then Stop: Exit Function
    MasterCatUpdateAction = CLng(MasterCatActionStr)

    '   Wip Dir
    '
    NewDir = UserProp_Get(oForm, pnWipDir)
    OldDir = UserProp_Get(oForm, pnWipDirSaved)

    Dim WipDirActionKey As String
    WipDirActionKey = _
        IIf(NewDir = OldDir, "1", "0") & _
        IIf(glbFSO.FolderExists(WipDirBasePath & NewDir), "1", "0") & _
        IIf(glbFSO.FolderExists(WipDirBasePath & OldDir), "1", "0")
        
    Dim WipDirActionStr As String
    If Not Tbl_TableConstFind(MasterCatActionTableConst, MasterCatActionKey, mccAction, WipDirActionStr) Then Stop: Exit Function
    WipDirUpdateAction = CLng(WipDirActionStr)

Return

End Function

' =====================================================================
'   Validate
' =====================================================================

Private Function CustForm_WipProjValidate(ByVal oForm As Outlook.PostItem, ByVal PropName As String) As Boolean
Const ThisProc = "CustForm_WipProjValidate"
CustForm_WipProjValidate = False

    Dim Cols() As String
    If Not Tbl_TableConstRow(FieldTableConst, PropName, Cols, ftcFieldName) Then Stop: Exit Function

    Dim oProp As Outlook.UserProperty
    Set oProp = oForm.UserProperties.Find(PropName, Cols(ftcPropType) = ptUser)
    If oProp Is Nothing Then Stop: Exit Function
    
    Dim PropValue As Variant
    PropValue = oProp.value

    '   If not required and null - Done
    '
    If Not CBool(Cols(ftcRequired)) And CStr(PropValue) = "" Then
        CustForm_WipProjValidate = True
        Exit Function
    End If

    GoSub Validate_Required
    GoSub Validate_MinLen
    GoSub Validate_MaxLen
    
CustForm_WipProjValidate = True
Exit Function

'   Required
'
Validate_Required:

    If Not CBool(Cols(ftcRequired)) Then Return
    If Not CStr(PropValue) = "" Then Return

    Msg_Box Proc:=ThisProc, Step:="Required?", _
            Text:="Field '" & Cols(ftcDisplayName) & "' is required."
    GoSub Validate_Focus
    Exit Function

Return

'   Minimum Lenght
'
Validate_MinLen:

    If Cols(ftcMin) = "" Then Return
    If Not Len(PropValue) < CLng(Cols(ftcMin)) Then Return
    
    Msg_Box Proc:=ThisProc, Step:="Minimum Length?", _
            Text:="Field '" & Cols(ftcDisplayName) & "' must be at least " & Cols(ftcMin) & " chars."
    GoSub Validate_Focus
    Exit Function

Return

'   Maximum Length
'
Validate_MaxLen:

    If Cols(ftcMax) = "" Then Return
    If Not Len(PropValue) > CLng(Cols(ftcMax)) Then Return
    
    Msg_Box Proc:=ThisProc, Step:="Maximum Length?", _
            Text:="Field '" & Cols(ftcDisplayName) & "' can not be more than " & Cols(ftcMax) & " chars."
    GoSub Validate_Focus
    Exit Function

Return

'   Set Focus
'
Validate_Focus:

    Dim oControl As MSForms.Control
    If Not CustForm_ControlByName(oForm, PropName, oControl) Then Stop: Exit Function
    oControl.SetFocus

Return

End Function

' =====================================================================
'   Sequence
' =====================================================================

Private Function CustForm_WipProjSeq(ByVal oForm As Outlook.PostItem) As Boolean
CustForm_WipProjSeq = False
    
    If Not Misc_RegExist(SeqRegKey) Then glbWshShell.RegWrite SeqRegKey, CStr(SeqInit), "REG_SZ"
    
    Dim SeqNum As Long
    SeqNum = CLng(glbWshShell.RegRead(SeqRegKey))
    If Not (SeqNum < SeqMax) Then Stop: Exit Function
    
    SeqNum = SeqNum + 1
    glbWshShell.RegWrite SeqRegKey, CStr(SeqNum), "REG_SZ"
    UserProp_Set oForm, pnSeq, CStr(SeqNum)
    
CustForm_WipProjSeq = True
End Function

' =====================================================================
'   Recalc
' =====================================================================

Private Function CustForm_WipProjRecalc(ByVal oForm As Outlook.PostItem) As Boolean
CustForm_WipProjRecalc = False

    Dim sTitle As String: sTitle = UserProp_Get(oForm, pnTitle)
    Dim sType As String: sType = UserProp_Get(oForm, pnType)
    Dim sSeq As String: sSeq = UserProp_Get(oForm, pnSeq)
    
    '   ID
    '
    Dim NewID As String
    NewID = sType & sSeq
    If Not UserProp_Get(oForm, pnID) = NewID Then
        UserProp_Set oForm, pnID, NewID
    End If
    
    '   Subject
    '
    Dim NewSubject As String
    NewSubject = sTitle & " (" & NewID & ")"
    If Not oForm.Subject = NewSubject Then
        oForm.Subject = NewSubject
    End If
    
    '   Master Cat
    '
    Dim NewMasterCat As String
    NewMasterCat = WIPCatPrefix & LCase(sType) & " | " & sTitle & " (" & NewID & ")"
    If Not UserProp_Get(oForm, pnMasterCat) = NewMasterCat Then
        UserProp_Set oForm, pnMasterCat, NewMasterCat
    End If
    
    '   WIP Dir
    '
    Dim NewWipDir As String
    NewWipDir = sType & "\" & sTitle & " (" & NewID & ")"
    If Not UserProp_Get(oForm, pnWipDir) = NewWipDir Then
        UserProp_Set oForm, pnWipDir, NewWipDir
    End If

CustForm_WipProjRecalc = True
End Function

' =====================================================================
'   New WIP Project
' =====================================================================

'   Create a new WIP Project Item in Projects & Show It
'
Public Sub CustForm_WipProjNew()

    Dim oFolder As Outlook.Folder
    Set oFolder = Folders_KnownPath(glbKnownPath_Projects)
    
    Dim oNewItem As Outlook.PostItem
    Set oNewItem = oFolder.Items.Add(glbCustForm_WipProject)
    
    If Inspector_ItemInspectorExist(oNewItem) Then Stop: Exit Sub
    oNewItem.GetInspector.Activate
    
End Sub

