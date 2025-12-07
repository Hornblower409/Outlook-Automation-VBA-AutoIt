VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppTimers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private TimersCollection As VBA.Collection
'

'   Property Get - Timer by TimerID or Name
'
Public Property Get Timers(ByVal Index As Variant) As Timer
    
    Set Timers = Nothing
    
    Select Case VarType(Index)
    Case vbString
        Dim Name As String: Name = CStr(Index)
        If InCollection(Name) Then Set Timers = TimersCollection.Item(Name)
    Case vbLong
        Dim TimerId As Long: TimerId = CLng(Index)
        For Each Timers In TimersCollection
            If Timers.TimerId = TimerId Then Exit Property
        Next Timers
    End Select
    
End Property

'   Collection Add
'
Public Sub Add(ByVal Timer As Timer)
    
    Remove Timer
    TimersCollection.Add Timer, Timer.Name
    
End Sub

'   Collection Remove
'
Public Sub Remove(ByVal Timer As Timer)

    If Not InCollection(Timer.Name) Then Exit Sub
    
    Dim OldTimer As Timer
    Set OldTimer = TimersCollection.Item(Timer.Name)
    OldTimer.Disable
    TimersCollection.Remove OldTimer.Name
    
End Sub

'   Init
'
Public Function Initialize() As Boolean
Initialize = False

    Set TimersCollection = New VBA.Collection
    
Initialize = True
End Function

'   Terminate
'
Private Sub Class_Terminate():  RemoveAll: End Sub
Public Sub RemoveAll()

    Dim Timer As Timer
    For Each Timer In TimersCollection
        Remove Timer
    Next Timer

End Sub

'   Is a Name in the Collection?
'
Private Function InCollection(ByVal Name As String) As Boolean

    On Error Resume Next
        TimersCollection.Item Name
        InCollection = (Err.Number = glbError_None)
    On Error GoTo 0

End Function
