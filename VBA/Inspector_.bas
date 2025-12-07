Attribute VB_Name = "Inspector_"
Option Explicit
Option Private Module

'   Inspector Info Record
'
Public Type Inspector_InspectorRec

    oItem           As Object               '   Item object
    oInspector      As Outlook.Inspector    '   Inspector object
    oInspShadow     As InspShadow           '   Inspector Shadow object
    InspectorIsNew  As Boolean              '   The Inspector was created by Inspector_RecFromItem
    
End Type

' =====================================================================
'   Inspector Helpers
'
'   ! CAUTION !
'
'       Inspector Objects don't die when they go out of scope becuase
'       they have been added to the Application.Inspectors collection.
'
'       You have to Close them or they will hang around and get used
'       again on the next Item.GetInspector call. But the Item they
'       reference may be gone or have changed.
'
'       An Inspector object does not have a default property. The only
'       way to access one from an Inspectors collection is by numeric
'       Index or For Each.
'
' =====================================================================

'   Populate an InspectorRec for an Item
'
Public Function Inspector_RecFromItem(ByVal oItem As Object, ByRef InspectorRec As Inspector_InspectorRec) As Boolean
Inspector_RecFromItem = False

    If oItem Is Nothing Then Stop: Exit Function

    With InspectorRec
    
        '   Get any existing Inspector
        '   If no existing Inspector - Get a new one.
        '
        If Not Inspector_RecInspectorFromItem(oItem, InspectorRec) Then Stop: Exit Function
        If .oInspector Is Nothing Then
        
            .InspectorIsNew = True
            Set .oInspector = .oItem.GetInspector
            DoEvents
            If .oInspector Is Nothing Then Stop: Exit Function
        
        End If
        
        '   Get the Inspector Shadow
        '
        Set .oInspShadow = glbAppShadows.InspShadows(.oInspector)
        If .oInspShadow Is Nothing Then Stop: Exit Function
    
    End With

Inspector_RecFromItem = True
End Function

'   Get any existing Inspector for an Item
'   If none - InspectorRec.oInspector = Nothing
'
Public Function Inspector_RecInspectorFromItem(ByVal oItem As Object, ByRef InspectorRec As Inspector_InspectorRec) As Boolean
Inspector_RecInspectorFromItem = False

    If oItem Is Nothing Then Stop: Exit Function
    Set InspectorRec.oItem = oItem
    Set InspectorRec.oInspector = Nothing
    
    '   If an Inspector is in the process of Closing while this code runs
    '   can cause all sorts of weird errors. Hence the Error Trap.
    '
    Dim oInspector As Outlook.Inspector
    On Error Resume Next
    
        For Each oInspector In Application.Inspectors: Do
                        
            If Not Err.Number = glbError_None Then
                Err.Clear
                Exit Do ' Next oInspector
            End If
                    
            If oInspector.CurrentItem Is InspectorRec.oItem Then
                Set InspectorRec.oInspector = oInspector
                Exit For
            End If
            
            If Not Err.Number = glbError_None Then
                Err.Clear
                Exit Do ' Next oInspector
            End If
            
        Loop While False: Next oInspector
    
    On Error GoTo 0

Inspector_RecInspectorFromItem = True
End Function

Public Function Inspector_ItemInspectorExist(ByVal oItem As Object) As Boolean
Inspector_ItemInspectorExist = False

    Dim InspectorRec As Inspector_InspectorRec
    If Not Inspector_RecInspectorFromItem(oItem, InspectorRec) Then Stop: Exit Function
    If InspectorRec.oInspector Is Nothing Then Exit Function

Inspector_ItemInspectorExist = True
End Function

Public Function Inspector_RecInspectorCloseIfNew(ByRef InspectorRec As Inspector_InspectorRec) As Boolean
Inspector_RecInspectorCloseIfNew = True

    With InspectorRec
    
        If .oInspector Is Nothing Then Exit Function
        If Not .InspectorIsNew Then Exit Function
        If .oInspector.CurrentItem Is Nothing Then Exit Function
        
        .oInspector.Close Outlook.OlInspectorClose.olDiscard
    
    End With
    
End Function
