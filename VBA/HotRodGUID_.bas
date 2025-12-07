Attribute VB_Name = "HotRodGUID_"
Option Explicit
Option Private Module

'   HotRodGUID Group User Props Record
'
Public Type HotRodGUID_HotRodGUIDRec

    oItem As Object
    GUID As String
    EntryId As String
    EntryIdMod As String
    
End Type

' =====================================================================
'   HotRodGUID Group User Props
' =====================================================================

'   Get HotRodGUID UserProps from an Item.
'   If any UserProp doesn't exist - Returns FALSE.
'
Public Function HotRodGUID_Get(ByVal oItem As Object, ByRef HotRodGUIDRec As HotRodGUID_HotRodGUIDRec) As Boolean
HotRodGUID_Get = True

    With HotRodGUIDRec
    
        Set .oItem = oItem
        If Not UserProp_GetPropTag(oItem, glbUserPropTag_HotRodGUID, .GUID) Then HotRodGUID_Get = False
        If Not UserProp_GetPropTag(oItem, glbUserPropTag_HotRodEntryId, .EntryId) Then HotRodGUID_Get = False
        If Not UserProp_GetPropTag(oItem, glbUserPropTag_HotRodEntryIdMod, .EntryIdMod) Then HotRodGUID_Get = False
        
    End With

End Function

'   Add HotRodGUID UserProps to an Item.
'
Public Function HotRodGUID_Add(ByRef HotRodGUIDRec As HotRodGUID_HotRodGUIDRec) As Boolean
HotRodGUID_Add = False

    With HotRodGUIDRec
    
        '   Generate new Values and Set the Props
        '
        If Not HotRodGUID_New(HotRodGUIDRec) Then Stop: Exit Function
        If Not HotRodGUID_Set(HotRodGUIDRec) Then Stop: Exit Function
        
    End With
    
HotRodGUID_Add = True
End Function

'   Generate new HotRodGUID UserProp Values.
'
Public Function HotRodGUID_New(ByRef HotRodGUIDRec As HotRodGUID_HotRodGUIDRec) As Boolean
HotRodGUID_New = False

    With HotRodGUIDRec
    
        .GUID = Misc_MakeGUID()
        .EntryId = .oItem.EntryId
        .EntryIdMod = Misc_NowString()
    
    End With
    
HotRodGUID_New = True
End Function

'   Set the HotRodGUID UserProps on an Item.
'
Public Function HotRodGUID_Set(ByRef HotRodGUIDRec As HotRodGUID_HotRodGUIDRec) As Boolean
HotRodGUID_Set = False

    With HotRodGUIDRec

        '   Can not operate on an unsaved Item
        '
        If Not .oItem.Saved Then Stop: Exit Function

        If Not UserProp_SetPropTag(.oItem, glbUserPropTag_HotRodGUID, .GUID) Then Exit Function
        If Not UserProp_SetPropTag(.oItem, glbUserPropTag_HotRodEntryId, .EntryId) Then Exit Function
        If Not UserProp_SetPropTag(.oItem, glbUserPropTag_HotRodEntryIdMod, .EntryIdMod) Then Exit Function
        
    End With

HotRodGUID_Set = True
End Function
