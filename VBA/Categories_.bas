Attribute VB_Name = "Categories_"
Option Explicit
Option Private Module

'   Assign a fixed set of Categories
'   Returns FALSE if anything goes wrong
'
'   !! Over writes any existing Cats !!
'
Public Function Categories_AssignFixed(ByVal Cats As String) As Boolean
Const ThisProc = "Categories_AssignFixed"

    Categories_AssignFixed = True
    
    If TypeOf ActiveWindow Is Outlook.Inspector Then
        ActiveInspector.CurrentItem.Categories = Cats
        Exit Function
    End If
    
    If TypeOf ActiveWindow Is Outlook.Explorer Then
    
        ' Get the current Explorer Selection
        '
        Dim Sel As Outlook.Selection
        Set Sel = ActiveExplorer.Selection
        
        ' If no items selected - exit
        '
        If Sel.count = 0 Then
            Categories_AssignFixed = False
            Exit Function
        End If
     
        ' For each selected Item replace any existing Cats and Save it
        '
        Dim Inx As Long
        Dim SelItem As Object
        For Inx = 1 To Sel.count
            Set SelItem = Sel.Item(Inx)
            SelItem.Categories = Cats
            SelItem.Save
        Next Inx
        
        Exit Function
        
    End If
    
    Msg_Box Proc:=ThisProc, Text:="That's interesting. ActiveWindow is not an Inspector and not an Explorer."
    Categories_AssignFixed = False
    
End Function

'   Show the All Categories selection dialog and let them pick
'
Public Function Categories_Assign() As Boolean
Const ThisProc = "Categories_Assign"

    '   Branch based on Inspector, Explorer, IMAP Explorer
    '
    If TypeOf ActiveWindow Is Outlook.Inspector Then
        Categories_Assign = Categories_Assign_Inspector()
    ElseIf TypeOf ActiveWindow Is Outlook.Explorer Then
        If IMAP_FolderIsIMAP(ActiveExplorer.CurrentFolder) Then
            Categories_Assign = Categories_Assign_Explorer()
        Else
            Categories_Assign = Categories_ShowAllCatsDialog()
        End If
    End If
    
End Function

'   Show the All Categories selection dialog for an Inspector
'
Private Function Categories_Assign_Inspector() As Boolean
Const ThisProc = "Categories_Assign_Inspector"
Categories_Assign_Inspector = False

    ' Show the All Cats dialog.
    '
    '   Don't need Save. ShowCategoriesDialog does not change Saved
    '
    Categories_ShowCategoriesDialog ActiveInspector.CurrentItem
    If ActiveInspector.CurrentItem.Categories = "" Then Exit Function
    
Categories_Assign_Inspector = True
End Function

'   Show the All Categories selection dialog for a NON IMAP Explorer
'
Private Function Categories_ShowAllCatsDialog() As Boolean
Const ThisProc = "Categories_ShowAllCatsDialog"
Categories_ShowAllCatsDialog = False
    
    '   Run my AutoIt script to resize the Name column when the Cat Dialog window opens
    '
    Utility_ShellExecute Application:=glbHotRodLnks & glbHotRodLnks_CatDialogResize, WindowMode:=0
        
    '   Show the All Categories (Master Cats) Dialog
    '
    Ribbon_ExecuteMSO ActiveExplorer, glbidMSO_AllCategories
    
    '   If no Cats Assigned - Exit
    '
    Dim Sel As Outlook.Selection
    Set Sel = ActiveExplorer.Selection
    If Sel.count = 0 Then Exit Function
    If Sel.Item(1).Categories = "" Then Exit Function
      
Categories_ShowAllCatsDialog = True
End Function

'   Show the All Categories selection dialog for an IMAP Explorer
'
'   !! Over writes any existing Cats with what is selected in the Dialog !!
'
Private Function Categories_Assign_Explorer() As Boolean
Const ThisProc = "Categories_Assign_Explorer"
Categories_Assign_Explorer = True

    ' Get the current Explorer Selection
    Dim Sel As Outlook.Selection
    Set Sel = ActiveExplorer.Selection
 
    ' If no items selected - exit
    If Sel.count = 0 Then
        Categories_Assign_Explorer = False
        Exit Function
    End If
 
    ' Open the All Cats dialog on the first item
    ' And pick off the Cats selected
    '
    '   No need to Save. Stupid cheats and doesn't set Saved = False
    '   even though the Item has been changed.
    '
    Dim selCats As String
    If Not Categories_ShowCategoriesDialog(Sel.Item(1)) Then
        Categories_Assign_Explorer = False
        Exit Function
    End If
    selCats = Sel.Item(1).Categories
    If selCats = "" Then
        Categories_Assign_Explorer = False
        Exit Function
    End If
    
    ' For each additional Item replace any
    ' existing Cats if different and Save it.
        
    Dim Inx As Long
    Dim SelItem As Object
    For Inx = 2 To Sel.count
        Set SelItem = Sel.Item(Inx)
        If SelItem.Categories <> selCats Then
            SelItem.Categories = selCats
            SelItem.Save
        End If
    Next Inx
     
End Function

'   Show the All Cats dialog for an Item
'
'       No need to Save. Stupid cheats and doesn't set Saved = False
'       even though the Item has been changed.
'
Public Function Categories_ShowCategoriesDialog(ByVal Item As Object) As Boolean
Const ThisProc = "Categories_ShowCategoriesDialog"
Categories_ShowCategoriesDialog = False

    '   Run my AutoIt script to resize the the Cat Dialog when it opens
    '
    Utility_ShellExecute Application:=glbHotRodLnks & glbHotRodLnks_CatDialogResize, WindowMode:=0
    
    '   Show the Dialog
    '
    On Error Resume Next
    
        Item.ShowCategoriesDialog
        Select Case Err.Number
            Case glbError_None
                '   Continue
            Case glbError_DoNotHavePermissions
                Msg_Box oErr:=Err, Icon:=vbExclamation, Proc:=ThisProc, Step:="ShowCategoriesDialog", Subject:=Item.Subject, _
                        Text:="Item is (probably) Read Only."
                Exit Function
            Case glbError_OpenRecurringFromInstance
                Msg_Box oErr:=Err, Icon:=vbExclamation, Proc:=ThisProc, Step:="ShowCategoriesDialog", Subject:=Item.Subject, _
                        Text:="Cat cannot be changed for a single instance of a recurring meeting."
                Exit Function
            Case Else
                Stop: Exit Function
        End Select
        
    On Error GoTo 0
    
    If Item.Categories = "" Then Exit Function

Categories_ShowCategoriesDialog = True
End Function

' =====================================================================
'   Categories Filter - Build a DASL Filter to EXACT MATCH Cats
' =====================================================================

Public Function Categories_CatFilter(ByVal Cats As String, ByRef Filter As String) As Boolean
Categories_CatFilter = False

    '   Setup
    '
    Dim urn As String
    urn = glbQuote & glbPropTag_Categories & glbQuote
    Const sOR As String = " OR "
    Filter = ""
    
    '   For each CAT - search for it in any position in the Cats String
    '
    '       Alone       Start       Middle          End
    '       "CAT"       "CAT, "     ", CAT, "       ", CAT"
    '
    Dim Cat As Variant
    For Each Cat In Split(Cats, ", ")
            
        Filter = Filter & sOR & _
                      urn & " LIKE '" & Cat & "'" & _
                sOR & urn & " LIKE '" & Cat & ", %'" & _
                sOR & urn & " LIKE '%, " & Cat & ", %'" & _
                sOR & urn & " LIKE '%, " & Cat & "'"
        
    Next Cat
    Filter = Mid(Filter, Len(sOR) + 1)

Categories_CatFilter = True
End Function

' =====================================================================
'   Categories Search - Open a new Explorer with selected Cats
' =====================================================================

Public Function Categories_CatSearch() As Boolean
Const ThisProc = "Categories_CatSearch"
Categories_CatSearch = False

    '   Create new Temp Item in Drafts
    '
    Dim Temp As Outlook.MailItem
    Set Temp = CreateItem(Outlook.olMailItem)
    
    '   Open the All CatString Dialog on it
    '   Pick off selected CatString
    '   Get rid of the Temp Item
    '   If no CatString selected - done
    '
    Categories_ShowCategoriesDialog Temp
    Dim CatString As String
    CatString = Temp.Categories
    Temp.Delete
    If CatString = "" Then Exit Function
    
    '   Get the Projects Primary View
    '
    Dim Views As Outlook.Views
    Set Views = Folders_KnownPath(glbKnownPath_Projects).Views
    
    Dim PrimaryView As Outlook.TableView
    Set PrimaryView = Views.Item(glbProjects_PrimaryViewName)
    If PrimaryView Is Nothing Then
        Msg_Box Proc:=ThisProc, Step:="Find Projects Primary View", _
                Text:="Projects Primary View '" & glbProjects_PrimaryViewName & "' does not exsist."
        Exit Function
    End If
    
    '   Create a new Cat Search View from the Projects Primary View
    '
    Dim TempView As Outlook.TableView
    Set TempView = Views.Item(glbProjects_CatSearchViewName)
    If Not TempView Is Nothing Then TempView.Delete
    Set TempView = PrimaryView.Copy(Name:=glbProjects_CatSearchViewName, SaveOption:=olViewSaveOptionThisFolderOnlyMe)
    
    '   Build a DASL Filter for the selected Cats
    '
    Dim Filter As String
    If Not Categories_CatFilter(CatString, Filter) Then Stop: Exit Function
    
    '   Apply the Filter
    '   Save the new Cat Search View
    '
    TempView.Filter = Filter
    TempView.Save
    
    '   Create a new Projects Explorer
    '   Apply the new Cat Search View
    '   Show the new Explorer
    '
    Dim TempExplorer As Outlook.Explorer
    Set TempExplorer = Folders_KnownPath(glbKnownPath_Projects).GetExplorer
    TempExplorer.CurrentView = glbProjects_CatSearchViewName
    TempExplorer.Activate
    
Categories_CatSearch = True
End Function

' =====================================================================
'   Find, Add, Remove Categories
' =====================================================================

'   Find a Cat
'
Public Function Categories_FindCat(ByVal Item As Object, ByVal Cat As String) As Boolean

    '   Look for ", Cat," in ", CATS,"
    '
    Categories_FindCat = (InStr(glbCatSep & " " & Item.Categories & glbCatSep, glbCatSep & " " & Cat & glbCatSep) > 0)

End Function

'   Add a Cat to an Item
'

Public Sub Categories_AddCat(ByVal Item As Object, ByVal Cat As String)

    If Categories_FindCat(Item, Cat) Then Exit Sub
    If Item.Categories = "" Then Item.Categories = Cat Else Item.Categories = Item.Categories & glbCatSep & Cat

End Sub

Public Sub Categories_AddCats(ByVal Item As Object, ByVal CatsList As String)

    Dim Cats() As String
    Cats = Split(CatsList, glbCatSep)
    
    Dim CatsIx As Long
    For CatsIx = 0 To UBound(Cats)
        Categories_AddCat Item, Trim(Cats(CatsIx))
    Next CatsIx

End Sub

'   Remove a Cat
'

Public Sub Categories_RemoveCat(ByVal Item As Object, ByVal Cat As String)

    Item.Categories = Categories_RemoveCatList(Item.Categories, Cat)
    
End Sub

Public Function Categories_RemoveCatList(ByVal CatsList As String, ByVal Cat As String) As String

    Dim CatArray() As String
    CatArray = Split(CatsList, glbCatSep)
    
    Dim Inx As Integer
    For Inx = 0 To UBound(CatArray)
        If Trim(CatArray(Inx)) = Cat Then CatArray(Inx) = ""
    Next Inx
    
    Categories_RemoveCatList = Join(Misc_ArrayCompress(CatArray), glbCatSep)
    
End Function


'   Remove Special Cats
'
'   (See glbSpecialCatPrefixTable)
'

Public Sub Categories_RemovePriorityCats(ByVal Item As Object)

    Item.Categories = Categories_RemovePriorityCatsList(Item.Categories)
    
End Sub

Public Function Categories_RemovePriorityCatsList(ByVal CatsList As String) As String

    Categories_RemovePriorityCatsList = Categories_RemovePrefixCatsList(CatsList, glbCatPrefixPriority)
        
End Function

Public Sub Categories_RemoveSpecialCats( _
    ByVal Item As Object, _
    Optional ByVal PriorityCats As Boolean = True _
    )

    Item.Categories = Categories_RemoveSpecialCatsList(Item.Categories, PriorityCats)
    
End Sub

Public Function Categories_RemoveSpecialCatsList( _
    ByVal CatsList As String, _
    Optional ByVal PriorityCats As Boolean = True _
    ) As String

    Dim NoPrefixes As String: NoPrefixes = CatsList

    '   Get a 1D Array of Special Cat Prefixes
    '
    Dim Prefixes() As String
    Prefixes = Tbl_TableConstList(glbSpecialCatPrefixTable)

    '   Walk the Prefixes List
    '
    Dim PrefixesInx As Long
    For PrefixesInx = LBound(Prefixes) To UBound(Prefixes): Do
    
        '   Skip Priority Cats if not called for
        '
        If Prefixes(PrefixesInx) = glbCatPrefixPriority Then
            If Not PriorityCats Then Exit Do    ' Next PrefixesInx
        End If
        
        '   Build a new NoPrefixes List without any Cats with that Prefix
        '
        NoPrefixes = Categories_RemovePrefixCatsList(NoPrefixes, Prefixes(PrefixesInx))
        
    Loop While False: Next PrefixesInx
    
    Categories_RemoveSpecialCatsList = NoPrefixes

End Function

Public Function Categories_RemovePrefixCatsList(ByVal CatsList As String, ByVal Prefix As String) As String

    '   Get a 1D Array of the Cats
    '
    Dim Cats() As String
    Cats = Split(CatsList, glbCatSep)
    
    '   Walk the Cats
    '   If the Cat starts with Prefix - empty it
    '
    Dim CatsInx As Long
    For CatsInx = LBound(Cats) To UBound(Cats)
        If InStr(1, Trim(Cats(CatsInx)), Prefix) = 1 Then Cats(CatsInx) = ""
    Next CatsInx
    
    '   Compress out any emptys and return the Cats List
    '
    Categories_RemovePrefixCatsList = Join(Misc_ArrayCompress(Cats), glbCatSep)

End Function

' =====================================================================
'   Master Cats
' =====================================================================

Public Function Categories_MasterCatsExist(ByVal Cat As String) As Boolean
Categories_MasterCatsExist = False

    '   Get the Categories collection
    '
    Dim oCats As Outlook.Categories
    Set oCats = Session.Categories
    If oCats.count < 1 Then Stop: Exit Function

    '   Does Cat exist?
    '   (Set oCat returns Nothing if not found. Does NOT throw an error!)
    '
    Dim oCat As Outlook.Category
    Set oCat = oCats.Item(Cat)
    If oCat Is Nothing Then Exit Function

Categories_MasterCatsExist = True
End Function

Public Function Categories_MasterCatsAssign(ByVal oItem As Object, ByVal Cat As String) As Boolean
Categories_MasterCatsAssign = False

    '   Cat Must Exist
    '
    If Not Categories_MasterCatsExist(Cat) Then Stop: Exit Function
    Categories_AddCat oItem, Cat
    
Categories_MasterCatsAssign = True
End Function

Public Function Categories_MasterCatsAdd(ByVal NewCat As String, ByVal Color As Outlook.OlCategoryColor) As Boolean
Categories_MasterCatsAdd = False

    '   NewCat Must Not Exist
    '
    If Categories_MasterCatsExist(NewCat) Then Stop: Exit Function
    
    '   Get the Categories collection
    '
    Dim oCats As Outlook.Categories
    Set oCats = Session.Categories
    If oCats.count < 1 Then Stop: Exit Function
    
    '   Add the Cat
    '
    Dim oCat As Outlook.Category
    Set oCat = oCats.Add(NewCat, Color)
    If oCat Is Nothing Then Stop: Exit Function
    
Categories_MasterCatsAdd = True
End Function

Public Function Categories_MasterCatsRename(ByVal OldCat As String, ByVal NewCat As String) As Boolean
Categories_MasterCatsRename = False

    '   OldCat Must Exist
    '
    If Not Categories_MasterCatsExist(OldCat) Then Stop: Exit Function
    
    '   NewCat Must NOT Exist
    '
    '   SPOS - If NewCat exist - Stupid just takes OldCat off the list.
    '
    If Categories_MasterCatsExist(NewCat) Then Stop: Exit Function
    
    '   Rename the Master Cat
    '   (After all I went through, this is all it takes)
    '
    '   Get the Categories collection
    '
    Dim oCats As Outlook.Categories
    Set oCats = Session.Categories
    If oCats.count < 1 Then Stop: Exit Function
    oCats.Item(OldCat).Name = NewCat
    
Categories_MasterCatsRename = True
End Function

Public Function Categories_MasterCatsRemove(ByVal Cat As String) As Boolean
Categories_MasterCatsRemove = False

    If Not Categories_MasterCatsExist(Cat) Then
        Categories_MasterCatsRemove = True
        Exit Function
    End If
    
    '   Remove the Master Cat
    '
    Dim oCats As Outlook.Categories
    Set oCats = Session.Categories
    oCats.Remove (Cat)
    
Categories_MasterCatsRemove = True
End Function


'   Backup the Master Cats List and keep the last NNN copies
'
'   (Called from Application_Startup)
'
Public Function Categories_MasterCatsBackup() As Boolean
Categories_MasterCatsBackup = False

    '   Get the Categories collection
    '
    Dim oCategories As Outlook.Categories
    Set oCategories = Session.Categories
    If oCategories.count < 1 Then Stop: Exit Function

    '   For each Cat in Categories
    '
    Dim aCatRows() As String
    ReDim aCatRows(1 To oCategories.count)
    Dim Inx As Long
    For Inx = 1 To oCategories.count

        '   Build a {Tab} delimited row of Cat Props in aCatRows
        '
        With oCategories.Item(Inx)
        aCatRows(Inx) = _
            .CategoryBorderColor & vbTab & _
            .CategoryGradientBottomColor & vbTab & _
            .CategoryGradientTopColor & vbTab & _
            .CategoryID & vbTab & _
            .Color & vbTab & _
            .Name & vbTab & _
            .ShortcutKey
        End With

    Next Inx

    '   Build a FileSpec for the Backup
    '
    Dim FileSpec As String
    FileSpec = glbMasterCatsBackupFolder & "\" & Misc_NowStamp() & "_Categories.tsv"
    
    '   Write aCatRows as a vbNewLine delimited file
    '
    If Not File_WriteText(FileSpec, Join(aCatRows, vbNewLine)) Then Stop: Exit Function
    
    '   Keep the last glbMasterCatsBackupVersionsToKeep versions
    '
    If Not File_DeleteOldByCreated(glbMasterCatsBackupFolder, glbMasterCatsBackupVersionsToKeep) Then Exit Function

Categories_MasterCatsBackup = True
End Function



