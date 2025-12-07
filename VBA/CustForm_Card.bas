Attribute VB_Name = "CustForm_Card"
Option Explicit
Option Private Module

' =====================================================================
'   Open
' =====================================================================

Public Function CustForm_CardOpen(ByVal oForm As Outlook.PostItem) As Boolean
CustForm_CardOpen = False

    '   Hook my Command Buttons
    '
    If Not CustForm_CardHookCmdButtons(oForm) Then Stop: Exit Function

CustForm_CardOpen = True
End Function

' =====================================================================
'   Custom Action
' =====================================================================

Public Function CustForm_CardCustAction(ByVal oForm As Outlook.PostItem, ByVal Action As String) As Boolean
CustForm_CardCustAction = False

CustForm_CardCustAction = True
End Function

' =====================================================================
'   Prop Change
' =====================================================================

Public Function CustForm_CardPropChange(ByVal oForm As Outlook.PostItem, ByVal IsStandardProp As Boolean, ByVal PropName As String) As Boolean
CustForm_CardPropChange = False

    Select Case True
        Case IsStandardProp
            If Not CustForm_CardStdProp(oForm, PropName) Then Stop: Exit Function
        Case Else
            If Not CustForm_CardUserProp(oForm, PropName) Then Stop: Exit Function
    End Select

CustForm_CardPropChange = True
End Function

Private Function CustForm_CardStdProp(ByVal oForm As Outlook.PostItem, ByVal PropName As String) As Boolean
CustForm_CardStdProp = False

CustForm_CardStdProp = True
End Function

Private Function CustForm_CardUserProp(ByVal oForm As Outlook.PostItem, ByVal PropName As String) As Boolean
CustForm_CardUserProp = False

    Select Case PropName
        Case "Level-1", "Level-2", "Level-3", "CardTitle"
            GoSub SubjectRecalc
        Case Else
            '   Continue
    End Select
    
CustForm_CardUserProp = True
Exit Function

'   Subject Recalc
'
SubjectRecalc:
    
    Dim Title As String:    Title = UserProp_Get(oForm, "CardTitle")
    Dim Level_1 As String:  Level_1 = UserProp_Get(oForm, "Level-1")
    Dim Level_2 As String:  Level_2 = UserProp_Get(oForm, "Level-2")
    Dim Level_3 As String:  Level_3 = UserProp_Get(oForm, "Level-3")
    
    Dim NewSubject As String
    NewSubject = _
        IIf(Len(Title) = 0, oForm.Subject, _
            Title & _
            IIf(Level_3 = "", "", " | " & Level_3) & _
            IIf(Level_2 = "", "", " | " & Level_2) & _
            IIf(Level_1 = "", "", " | " & Level_1) _
        )

    If oForm.Subject <> NewSubject Then oForm.Subject = NewSubject

    Return
    
End Function

' =====================================================================
'   Command Buttons
' =====================================================================

Public Function CustForm_CardHookCmdButtons(ByVal oForm As Outlook.PostItem) As Boolean
CustForm_CardHookCmdButtons = False

    '   Get the Form's InspectorRec
    '
    Dim InspectorRec As Inspector_InspectorRec
    If Not Inspector_RecFromItem(oForm, InspectorRec) Then Stop: Exit Function
    
    With InspectorRec
    
        '   Get the "Clone Card" Control by Name and Hook it
        '
        Dim oControl As MSForms.Control
        If Not CustForm_ControlByName(oForm, "ButtonCloneCard", oControl) Then Stop: Exit Function
        .oInspShadow.CmdButtonHook oControl

    End With

CustForm_CardHookCmdButtons = True
End Function

Public Function CustForm_CardClickCmdButton(ByVal oForm As Outlook.PostItem, ByVal oCmdButton As MSForms.Control) As Boolean
CustForm_CardClickCmdButton = False

    Select Case oCmdButton.Name
        Case "ButtonCloneCard"
            If Not CustForm_CardClone(oForm) Then Stop: Exit Function
        Case Else
            Stop: Exit Function
    End Select

CustForm_CardClickCmdButton = True
End Function

' =====================================================================
'   Clone
' =====================================================================

Private Function CustForm_CardClone(ByVal oCard As Object) As Boolean
Const ThisProc = "CustForm_CardClone"
CustForm_CardClone = False

    '   Create a new Post Item of Card Class in the current Folder Items collection
    '
    Dim NewPost As Outlook.PostItem
    Set NewPost = oCard.Parent.Items.Add(glbCustForm_Card)
        
    '   Copy StdProps, that I want, from the old Card
    '
    NewPost.Subject = oCard.Subject
        
    '   Copy UserProps, that exist on the New Card and that I want, from the old Card
    '
    Dim oNewUserProp As Outlook.UserProperty
    For Each oNewUserProp In NewPost.UserProperties: Do
    
        '   Skip any UserProps I don't want copied
        '
        Select Case oNewUserProp.Name
        
            Case _
                 glbUserPropTag_HotRodGUID, _
                 glbUserPropTag_HotRodEntryId, _
                 glbUserPropTag_HotRodEntryIdMod
                 Exit Do ' Next Inx
            Case Else
                ' Continue
        End Select
        
        oNewUserProp.value = UserProp_Get(oCard, oNewUserProp.Name)
        
    Loop While False: Next oNewUserProp
    
    '   Show them the pretty new Clone
    '
    NewPost.GetInspector.Activate   ' 2025-02-19 Was "NewPost.Display"
    
CustForm_CardClone = True
End Function

