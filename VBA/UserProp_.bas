Attribute VB_Name = "UserProp_"
Option Explicit
Option Private Module

' =====================================================================
'   User Prop - UserProperties
' =====================================================================

Public Function UserProp_Get(ByVal oItem As Object, ByVal PropName As String) As Variant

    Dim oUserProp As Outlook.UserProperty
    Set oUserProp = oItem.UserProperties.Find(PropName, True)
    If oUserProp Is Nothing Then Stop: Exit Function
    
    On Error Resume Next
        UserProp_Get = oUserProp.value
        If Err.Number <> glbError_None Then Stop: Exit Function
    On Error GoTo 0
        
End Function

Public Function UserProp_Set(ByVal oItem As Object, ByVal PropName As String, ByVal value As Variant) As Boolean
UserProp_Set = False

    Dim oUserProp As Outlook.UserProperty
    Set oUserProp = oItem.UserProperties.Find(PropName, True)
    If oUserProp Is Nothing Then Stop: Exit Function

    On Error Resume Next
        oUserProp.value = value
        If Err.Number <> glbError_None Then Stop: Exit Function
    On Error GoTo 0
        
UserProp_Set = True
End Function

Public Function UserProp_Obj(ByVal oItem As Object, ByVal PropName As String) As Outlook.UserProperty

    Set UserProp_Obj = oItem.UserProperties.Find(PropName, True)
    If UserProp_Obj Is Nothing Then Stop: Exit Function

End Function

' =====================================================================
'   User Prop - PropTag
' =====================================================================

'   Get a UserProp Value By PropTag from an Item. If UserProp doesn't exist - Returns FALSE.
'
Public Function UserProp_GetPropTag( _
    ByVal oItem As Object, _
    ByVal UserPropTag As String, _
    ByRef UserPropValue As String _
    ) As Boolean
UserProp_GetPropTag = False

    UserPropValue = ""
    
    Dim UserPropSchema As String
    UserPropSchema = glbUserPropsSchemaPrefix & UserPropTag & glbUserPropsSchemaSuffix
    
    Dim PA As Outlook.PropertyAccessor
    Set PA = oItem.PropertyAccessor
    
    On Error Resume Next
        UserPropValue = PA.GetProperty(UserPropSchema)
        Select Case Err.Number
            Case glbError_None
            Case glbError_PAPropertyNotFound
                Exit Function
            Case Else
                Stop: Exit Function
        End Select
    On Error GoTo 0

UserProp_GetPropTag = True
End Function

'   Add/Set a UserProp Value By PropTag for an Item. If UserProp doesn't exist - Add it.
'
'   SPOS - Add/Set UserProp does NOT set Item.Saved = False. You must Save or the change is lost.
'
Public Function UserProp_SetPropTag( _
    ByVal oItem As Object, _
    ByVal UserPropTag As String, _
    ByVal UserPropValue As String _
    ) As Boolean
UserProp_SetPropTag = False
    
    '   Can not operate on an unsaved Item
    '
    If Not oItem.Saved Then Stop: Exit Function
    
    '   Set the UserProp. If UserProp doesn't exist - it is added automatically.
    '
    Dim UserPropSchema As String
    UserPropSchema = glbUserPropsSchemaPrefix & UserPropTag & glbUserPropsSchemaSuffix
    
    Dim PA As Outlook.PropertyAccessor
    Set PA = oItem.PropertyAccessor
    
    On Error Resume Next
        PA.SetProperty UserPropSchema, UserPropValue
        Select Case Err.Number
            Case glbError_None
            Case Else
                Stop: Exit Function
        End Select
    On Error GoTo 0
    
    oItem.Save
    
UserProp_SetPropTag = True
End Function

'   Delete a UserProp By PropTag from an Item if it exist.
'
'   SPOS - Delete UserProp does NOT set Item.Saved = False. You must Save or the change is lost.
'
'   !!!!!  Does NOT check for Saved or do a Save after Delete. It's up to the Caller  !!!!
'
Public Function UserProp_DeletePropTag( _
    ByVal oItem As Object, _
    ByVal UserPropTag As String _
    ) As Boolean
UserProp_DeletePropTag = False
    
    Dim UserPropSchema As String
    UserPropSchema = glbUserPropsSchemaPrefix & UserPropTag & glbUserPropsSchemaSuffix
    
    Dim PA As Outlook.PropertyAccessor
    Set PA = oItem.PropertyAccessor
    
    On Error Resume Next
        PA.DeleteProperty UserPropSchema
        Select Case Err.Number
            Case glbError_None
                '   Continue
            Case glbError_PAPropertyNotFound
                '   Continue
            Case glbError_OpNotSupported
                Exit Function
            Case Else
                Stop: Exit Function
        End Select
    On Error GoTo 0
    
UserProp_DeletePropTag = True
End Function

' =====================================================================
'   User Props - Filter
' =====================================================================

'   Filter (Search/Find/Restrict) a Folder for a UserProp Value
'
'       If you are trying to use the Find or Restrict methods with user-defined fields,
'       the fields must be defined in the folder, otherwise an error will occur.
'
Public Function UserProp_FilterPropTag( _
    ByVal FolderPath As String, _
    ByVal UserPropTag As String, _
    ByVal UserPropValue As String, _
    ByRef oResults As Outlook.Items _
    ) As Boolean
UserProp_FilterPropTag = False

    '   Get a Folder
    '
    Dim oFolder As Outlook.Folder
    Set oFolder = Folders_Path(FolderPath)
    If oFolder Is Nothing Then Stop: Exit Function
   
    '   If the UserPropTag is not defined in the Folder - add it
    '
    Dim oUserDefProps As Outlook.UserDefinedProperties
    Set oUserDefProps = oFolder.UserDefinedProperties
    Dim oUserDefProp As Outlook.UserDefinedProperty
    Set oUserDefProp = oUserDefProps.Find(UserPropTag)
    If oUserDefProp Is Nothing Then Set oUserDefProp = oUserDefProps.Add(Name:=UserPropTag, Type:=Outlook.OlUserPropertyType.olText)
        
    '   Construct the filter
    '   Get the Results Items collection
    '
    Dim sFilter As String
    sFilter = "[" & UserPropTag & "] = " & glbQuote & UserPropValue & glbQuote
    Set oResults = oFolder.Items.Restrict(sFilter)
    If oResults.count < 1 Then Exit Function

UserProp_FilterPropTag = True
End Function

' =====================================================================
'   User Props - Manual
' =====================================================================

'   Delete a User Prop from an Item by PropTag or UserProperties.Remove
'
'   !!  Does NOT save the Item. You must do it manually !!
'
'   To delete a User Prop from a Custom Form:
'
'       Design the Form
'       Run this Sub
'       Publish the Form
'       Close the Form
'       Clear Forms Cache
'
Public Sub UserProp_DeleteItem()

    ' *****************************
        Dim UserPropTag As String
        UserPropTag = "Subject"
    ' *****************************

    Dim oItem As Object
    If Not Misc_GetActiveItem(oItem) Then Stop: Exit Sub

    '   Try Delete using UserProperties
    '
    Dim oUserProps As Outlook.UserProperties
    Set oUserProps = oItem.UserProperties
    
    Dim oUserProp As Outlook.UserProperty
    Dim UserPropsInx As Long
    For UserPropsInx = 1 To oUserProps.count
    
        Set oUserProp = oUserProps.Item(UserPropsInx)
        If oUserProp.Name = UserPropTag Then
            oUserProps.Remove UserPropsInx
            Stop
            Exit Sub
        End If
    
    Next UserPropsInx

    '   Try Delete using a PropTag
    '
    Dim UserPropValue As String
    If UserProp_GetPropTag(oItem, UserPropTag, UserPropValue) Then
    
        If Not UserProp_DeletePropTag(oItem, UserPropTag) Then Stop: Exit Sub
        Stop
        Exit Sub
        
    End If
    
    '   Both Failed
    '
    Stop
    
End Sub

'   Delete a User Prop By PropTag from all Items in a Folder
'
Public Sub UserProp_DeleteFolderPropTag()
    
    ' **************************************************************
        Dim FolderPath As String:   FolderPath = "\\Cards\Cards"
        Dim UserPropTag As String:  UserPropTag = glbUserPropTag_HotRodEntryIdMod
    ' **************************************************************
    
    '   Get a Folder
    '
    Dim oFolder As Outlook.Folder
    Set oFolder = Folders_Path(FolderPath)
    If oFolder Is Nothing Then Stop: Exit Sub
    
    '   Loop thru all Items
    '
    Dim ChangedCount As Long
    Dim LoopCount As Long
    Dim oItem As Object
    Dim UserPropValue As String
    For Each oItem In oFolder.Items
    
        LoopCount = LoopCount + 1
        If (LoopCount Mod 100) = 0 Then
            DoEvents
            Debug.Print LoopCount
        End If

        If UserProp_GetPropTag(oItem, UserPropTag, UserPropValue) Then
    
            '   SPOS - Delete UserProp does NOT set Item.Saved = False. But you must Save or the chage is lost.
            '
            If Not UserProp_DeletePropTag(oItem, UserPropTag) Then Stop: Exit Sub
            oItem.Save
            ChangedCount = ChangedCount + 1
            
        End If
        
    Next oItem
    
    Debug.Print "Changed: " & ChangedCount
    
End Sub



