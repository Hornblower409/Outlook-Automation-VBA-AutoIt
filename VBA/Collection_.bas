Attribute VB_Name = "Collection_"
Option Explicit
Option Private Module

' =====================================================================
'   VBA Collections
' =====================================================================

'   Does Collection.Item("Key") exist?
'
Public Function Collection_KeyExist(ByVal Collection As VBA.Collection, ByVal Key As String) As Boolean

    '   Sets Err.Number <> 0 if Key is not in Collection
    '
    On Error Resume Next
    
        Collection.Item Key
        Collection_KeyExist = (Err.Number = glbError_None)
        
    On Error GoTo 0

End Function

'   Get Item by Key and return Found as True/False
'
Public Function Collection_Get(ByVal Collection As VBA.Collection, ByVal Key As String, ByRef Item As Variant) As Boolean

    On Error Resume Next
    
        Item = Collection.Item(Key)
        Collection_Get = (Err.Number = glbError_None)
        
    On Error GoTo 0

End Function

'   Move an Outlook.Items Selection into a VBA.Collection
'
'       So I don't have to deal with the Outlook Selection changing as Items are changed/deleted.
'       See Card "Items.Restrict (SPOS)"
'
Public Function Collection_FromSelection(ByVal OutlookSelection As Outlook.Items) As VBA.Collection

    Dim VBACollection As VBA.Collection
    Set VBACollection = New VBA.Collection
    Dim Index As Long
    For Index = 1 To OutlookSelection.count
        VBACollection.Add OutlookSelection.Item(Index), CStr(Index)
    Next Index
    
    Set Collection_FromSelection = VBACollection

End Function

'   Get the Results of a Folder Restrict as a VBA.Collection
'
Public Function Collection_FromRestrict(ByVal SQLRestrict As String, ByVal Folder As Outlook.Folder, ByRef Results As VBA.Collection) As Boolean
Collection_FromRestrict = False

    Dim Selection As Outlook.Items
    Set Selection = Folder.Items.Restrict(SQLRestrict)
    Set Results = Collection_FromSelection(Selection)

Collection_FromRestrict = True
End Function
 
