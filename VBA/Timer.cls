VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Timer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
'
'   https://stackoverflow.com/questions/12257985/outlook-vba-run-a-code-every-half-an-hour
'   http://www.cpearson.com/excel/OnTime.aspx
'
'   Prototype: See EventX_ExitEventScopeTimer
'
' =====================================================================

Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal P1 As Long, ByVal P2 As Long, ByVal Period As Long, ByVal Callback As Long) As Long
Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal P1 As Long, ByVal TimerId As Long) As Long

Private myName As String
Private myPeriod As Long
Private myRetries As Long
Private myCallback As Long
Private myStashVar As Variant
Private myStashObj As Object
Private myTimerID As Long

Public Property Get Period() As Long:                           Period = myPeriod:                  End Property
Public Property Get Retries() As Long:                          Retries = myRetries:                End Property
Public Property Get StashVar() As Variant:                      StashVar = myStashVar:              End Property
Public Property Get StashObj() As Object:                       Set StashObj = myStashObj:          End Property
Public Property Get TimerId() As Long:                          TimerId = myTimerID:                End Property
Public Property Get Enabled() As Boolean:                       Enabled = Not (myTimerID = 0):      End Property

Public Property Get Name() As String:                           Name = myName:                      End Property
Public Property Let Name(ByVal newName As String)

    '   Must have a unique name for the AppTimers Collection
    '
    '   - Add a Time Stamp and GUID to the name
    '
    myName = newName & " " & Misc_NowStamp() & " " & Misc_MakeGUID()
    
End Property

Public Property Let Period(ByVal newPeriod As Long):            myPeriod = newPeriod:               End Property
Public Property Let Retries(ByVal newRetries As Long):          myRetries = newRetries:             End Property
Public Property Let StashVar(ByVal newStashVar As Variant):     myStashVar = newStashVar:           End Property
Public Property Set StashObj(ByVal newStashObj As Object):      Set myStashObj = newStashObj:       End Property        ' For Let StashObj use "Set .StashObj = Object"
Public Sub Callback(ByVal ProcAddress As Long):                 myCallback = ProcAddress:           End Sub             ' For Let Callback use ".Callback AddressOf ProcName"

Public Sub Enable()

    If myTimerID <> 0 Then Disable
    myTimerID = SetTimer(0, 0, myPeriod, myCallback)
    If myTimerID = 0 Then Stop: Exit Sub
    
End Sub

Public Sub Disable()

    If myTimerID = 0 Then Exit Sub
    KillTimer 0, myTimerID
    myTimerID = 0

End Sub

Public Function Retry() As Boolean

    myRetries = myRetries - 1
    Retry = Not (myRetries < 1)
    
End Function
