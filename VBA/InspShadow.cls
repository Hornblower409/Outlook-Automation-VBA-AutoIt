VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InspShadow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents MyInspector      As Outlook.Inspector
Attribute MyInspector.VB_VarHelpID = -1

Private WithEvents AppointmentItem  As Outlook.AppointmentItem
Attribute AppointmentItem.VB_VarHelpID = -1
Private WithEvents ContactItem      As Outlook.ContactItem
Attribute ContactItem.VB_VarHelpID = -1
Private WithEvents MailItem         As Outlook.MailItem
Attribute MailItem.VB_VarHelpID = -1
Private WithEvents MeetingItem      As Outlook.MeetingItem
Attribute MeetingItem.VB_VarHelpID = -1
Private WithEvents PostItem         As Outlook.PostItem
Attribute PostItem.VB_VarHelpID = -1
Private WithEvents TaskItem         As Outlook.TaskItem
Attribute TaskItem.VB_VarHelpID = -1

Private WithEvents CmdButton0       As MSForms.CommandButton
Attribute CmdButton0.VB_VarHelpID = -1
Private WithEvents CmdButton1       As MSForms.CommandButton
Attribute CmdButton1.VB_VarHelpID = -1
Private WithEvents CmdButton2       As MSForms.CommandButton
Attribute CmdButton2.VB_VarHelpID = -1
Private WithEvents CmdButton3       As MSForms.CommandButton
Attribute CmdButton3.VB_VarHelpID = -1
Private WithEvents CmdButton4       As MSForms.CommandButton
Attribute CmdButton4.VB_VarHelpID = -1

Private CmdButtons(5) As Boolean

Private myShadowKey As String                           ' This instance InspShadowKey
Private myCleanupResponseDone As Boolean                ' Cleanup been done on any Response Item for this Inspector?
Private myCleanupSkipOnSend As Boolean                  ' Skip automatic Cleanup on the next Send for this Inspector?
Private myCleanupResponseBypass As Boolean              ' Bypass Cleanup for any Response Item for this Inspector?
'

' =====================================================================
'                   Inspector Events
' =====================================================================

'   Inspector Init (Open)
'
Public Function Initialize(ByVal TheInspector As Outlook.Inspector, ByVal TheShadowKey As Long) As Boolean
Initialize = False

    '   Shadow Inspector Housekeeping
    '
    Set MyInspector = TheInspector
    myShadowKey = CStr(TheShadowKey)
    glbAppShadows.InspShadowsAdd Me, myShadowKey
    
    '   Hook my Item Event
    '
    With MyInspector
    
        Select Case True
        
            Case TypeOf .CurrentItem Is Outlook.AppointmentItem
                Set AppointmentItem = .CurrentItem
            Case TypeOf .CurrentItem Is Outlook.ContactItem
                Set ContactItem = .CurrentItem
            Case TypeOf .CurrentItem Is Outlook.MailItem
                Set MailItem = .CurrentItem
            Case TypeOf .CurrentItem Is Outlook.MeetingItem
                Set MeetingItem = .CurrentItem
            Case TypeOf .CurrentItem Is Outlook.PostItem
                Set PostItem = .CurrentItem
            Case TypeOf .CurrentItem Is Outlook.TaskItem
                Set TaskItem = .CurrentItem
            Case TypeOf .CurrentItem Is Outlook.NoteItem
                ' Sticky Note - No hooks.
            Case TypeOf .CurrentItem Is Outlook.DistListItem
                ' Contact Distribution List - No hooks.
            Case Else
                Stop: Exit Function
        End Select
        
    End With
    
Initialize = True
End Function

'   Inspector Close
'
Private Sub MyInspector_Close()

    If Not Event_CloseInspector(MyInspector.CurrentItem, Me) Then Exit Sub
    glbAppShadows.InspShadowsRemove myShadowKey

End Sub

'   Inspector Activate
'
Private Sub MyInspector_Activate()

    If Not Event_ActivateInspector(MyInspector.CurrentItem, Me) Then Exit Sub

End Sub

' =====================================================================
'                   CommandButtons
' =====================================================================

Public Sub CmdButtonHook(ByVal oCmdButton As MSForms.CommandButton)

    Dim CmdButtonFree As Long
    CmdButtonFree = -1
    
    Dim Index As Long
    For Index = 0 To UBound(CmdButtons)
        If Not CmdButtons(Index) Then
            CmdButtonFree = Index
            Exit For
        End If
    Next Index
    If CmdButtonFree < 0 Then Stop: Exit Sub
    
    Select Case CmdButtonFree
        Case 0: Set CmdButton0 = oCmdButton
        Case 1: Set CmdButton1 = oCmdButton
        Case 2: Set CmdButton2 = oCmdButton
        Case 3: Set CmdButton3 = oCmdButton
        Case 4: Set CmdButton4 = oCmdButton
        Case Else: Stop: Exit Sub
    End Select
    CmdButtons(CmdButtonFree) = True
    
End Sub

Private Sub CmdButton0_Click(): CmdButtonClick CmdButton0:  End Sub
Private Sub CmdButton1_Click(): CmdButtonClick CmdButton1:  End Sub
Private Sub CmdButton2_Click(): CmdButtonClick CmdButton2:  End Sub
Private Sub CmdButton3_Click(): CmdButtonClick CmdButton3:  End Sub
Private Sub CmdButton4_Click(): CmdButtonClick CmdButton4:  End Sub
Private Sub CmdButtonClick(ByVal oCmdButton As MSForms.CommandButton):  Event_ClickCmdButton MyInspector.CurrentItem, oCmdButton:                       End Sub


' =====================================================================
'                   Item Events
' =====================================================================

'   Open
'
Private Sub AppointmentItem_Open(ByRef Cancel As Boolean):                          Cancel = Not Event_OpenItemEvent(AppointmentItem):                              End Sub
Private Sub ContactItem_Open(ByRef Cancel As Boolean):                              Cancel = Not Event_OpenItemEvent(ContactItem):                                  End Sub
Private Sub MailItem_Open(ByRef Cancel As Boolean):                                 Cancel = Not Event_OpenItemEvent(MailItem):                                     End Sub
Private Sub PostItem_Open(ByRef Cancel As Boolean):                                 Cancel = Not Event_OpenItemEvent(PostItem):                                     End Sub
Private Sub MeetingItem_Open(ByRef Cancel As Boolean):                              Cancel = Not Event_OpenItemEvent(MeetingItem):                                  End Sub
Private Sub TaskItem_Open(ByRef Cancel As Boolean):                                 Cancel = Not Event_OpenItemEvent(TaskItem):                                     End Sub


'   Misc
'
Private Sub MailItem_BeforeDelete(ByVal Item As Object, ByRef Cancel As Boolean):   Cancel = Not Event_BeforeDelete(MailItem):                                      End Sub

'   Response
'
Private Sub MailItem_Reply(ByVal Response As Object, ByRef Cancel As Boolean):          Response_Main MailItem, Response, Cancel, glbResponse_Reply:                End Sub
Private Sub MeetingItem_Reply(ByVal Response As Object, ByRef Cancel As Boolean):       Response_Main MeetingItem, Response, Cancel, glbResponse_Reply:             End Sub
Private Sub PostItem_Reply(ByVal Response As Object, ByRef Cancel As Boolean):          Response_Main PostItem, Response, Cancel, glbResponse_Reply:                End Sub
Private Sub MailItem_ReplyAll(ByVal Response As Object, ByRef Cancel As Boolean):       Response_Main MailItem, Response, Cancel, glbResponse_ReplyAll:             End Sub
Private Sub MeetingItem_ReplyAll(ByVal Response As Object, ByRef Cancel As Boolean):    Response_Main MeetingItem, Response, Cancel, glbResponse_ReplyAll:          End Sub
Private Sub PostItem_ReplyAll(ByVal Response As Object, ByRef Cancel As Boolean):       Response_Main PostItem, Response, Cancel, glbResponse_ReplyAll:             End Sub
Private Sub MailItem_Forward(ByVal Response As Object, ByRef Cancel As Boolean):        Response_Main MailItem, Response, Cancel, glbResponse_Forward:              End Sub
Private Sub MeetingItem_Forward(ByVal Response As Object, ByRef Cancel As Boolean):     Response_Main MeetingItem, Response, Cancel, glbResponse_Forward:           End Sub
Private Sub PostItem_Forward(ByVal Response As Object, ByRef Cancel As Boolean):        Response_Main PostItem, Response, Cancel, glbResponse_Forward:              End Sub
'

'   Prop Change
'
Private Sub PostItem_PropertyChange(ByVal Name As String):                          Event_PropChangePost PostItem, True, Name:                                      End Sub
Private Sub PostItem_CustomPropertyChange(ByVal Name As String):                    Event_PropChangePost PostItem, False, Name:                                     End Sub

'   Write
'
Private Sub PostItem_Write(Cancel As Boolean):                                      Cancel = Not Event_WriteItem(PostItem):                                         End Sub

'   Custom Action
'
Private Sub PostItem_CustomAction(ByVal Action As Object, ByVal Response As Object, ByRef Cancel As Boolean):   Event_CustomActionPost Action, Response, Cancel:    End Sub

' =====================================================================
'               Property Accessors
' =====================================================================
'
Public Property Get Inspector() As Outlook.Inspector:                                       Set Inspector = MyInspector:                                            End Property
Public Property Get ShadowKey() As String:                                                  ShadowKey = myShadowKey:                                                End Property
Public Property Get CleanupSkipOnSend() As Boolean:                                         CleanupSkipOnSend = myCleanupSkipOnSend:                                End Property
Public Property Let CleanupSkipOnSend(ByVal CleanupSkipOnSend As Boolean):                  myCleanupSkipOnSend = CleanupSkipOnSend:                                End Property
Public Property Get CleanupResponseDone() As Boolean:                                       CleanupResponseDone = myCleanupResponseDone:                            End Property
Public Property Let CleanupResponseDone(ByVal CleanupResponseDone As Boolean):              myCleanupResponseDone = CleanupResponseDone:                            End Property
Public Property Get CleanupResponseBypass() As Boolean:                                     CleanupResponseBypass = myCleanupResponseBypass:                            End Property
Public Property Let CleanupResponseBypass(ByVal CleanupResponseBypass As Boolean):          myCleanupResponseBypass = CleanupResponseBypass:                            End Property
'
