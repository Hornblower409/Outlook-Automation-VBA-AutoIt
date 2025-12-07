VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FollowUp_Form 
   Caption         =   "Follow Up with Reminder"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6690
   OleObjectBlob   =   "FollowUp_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FollowUp_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Canceled As Boolean
Public Cleared As Boolean
Public Title As String
Public DateTime As Variant

Private ActivateInProgress As Boolean
Private InitialTitle As String
Private InitialDateTime As Variant
Private Const FormatDate As String = "ddd mmm dd yyyy"
Private Const FormatTime As String = "Hh:Nn AM/PM"

Private Sub UserForm_Activate()
ActivateInProgress = True

    InitialTitle = Title
    InitialDateTime = DateTime
    
    TextTitle.value = Title
    TextDate.value = Format(DateTime, FormatDate)
    TextTime.value = Format(DateTime, FormatTime)
    
    DatePicker.Caption = ChrW(glbUnicode_DownTriangle)
    
    If Title = "" Then ButtonClear.Enabled = False
        
    ButtonOK.Enabled = False
    If TextTitle.value <> "" Then Validate
    
ActivateInProgress = False
End Sub

Private Sub TextTitle_Change()

    If Not ActivateInProgress Then Validate

End Sub

Private Sub TextTitle_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Validate

End Sub

Private Sub TextDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If IsEmpty(ConvDate()) Then
        TextDate.value = Format(Date, FormatDate)
    Else
        TextDate.value = Format(ConvDate(), FormatDate)
    End If
    
    Validate
    
End Sub

Private Sub TextTime_Change()

    If Not ActivateInProgress Then Validate
    
End Sub

Private Sub TextTime_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If IsEmpty(ConvTime()) Then
        TextTime.value = Format(Now, FormatTime)
    Else
        TextTime.value = Format(ConvTime(), FormatTime)
    End If

    Validate

End Sub

Private Sub ButtonOK_Click()

    Me.Hide

End Sub

Private Sub ButtonCancel_Click()

    OnCancel
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub
Private Sub OnCancel()

    Canceled = True
    Hide

End Sub

Private Sub ButtonClear_Click()

    Cleared = True
    Me.Hide
    
End Sub

Private Sub ButtonReset_Click()
ActivateInProgress = True

    TextTitle.value = InitialTitle
    TextDate.value = Format(InitialDateTime, FormatDate)
    TextTime.value = Format(InitialDateTime, FormatTime)

    ButtonOK.Enabled = False
    If TextTitle.value <> "" Then Validate

ActivateInProgress = False
End Sub

Private Sub Validate()

    LabelError.Caption = ""
    
    ButtonOK.Enabled = True
    ButtonReset.Enabled = False
    
    If (TextTitle.value = InitialTitle) _
    And (TextDate.value = Format(InitialDateTime, FormatDate)) _
    And (TextTime.value = Format(InitialDateTime, FormatTime)) _
    Then Exit Sub
    
    ButtonOK.Enabled = False
    ButtonReset.Enabled = True
    
    If TextTitle.value = "" Then
        LabelError.Caption = "Title can not be blank."
        TextTitle.SetFocus
        Exit Sub
    End If
    
    If InStr(TextTitle.value, vbTab) > 0 Then
        LabelError.Caption = "Title can not contain Tabs."
        TextTitle.SetFocus
        Exit Sub
    End If
    
    If IsEmpty(ConvDate()) Then
        LabelError.Caption = "Invalid Reminder Date."
        TextDate.SetFocus
        Exit Sub
    End If
    
    If IsEmpty(ConvTime()) Then
        LabelError.Caption = "Invalid Reminder Time."
        TextTime.SetFocus
        Exit Sub
    End If
    
    '   Cast required
    '
    If CDate(ConvDate()) < Date Then
        LabelError.Caption = "Reminder Date occurs in the past."
        TextDate.SetFocus
        Exit Sub
    End If
        
    '   Need the two cast or it doesn't work
    '   (Stupid concats the values)
    '
    DateTime = CDate(ConvDate()) + CDate(ConvTime())
    
    '   Cast required
    '
    If CDate(DateTime) < Now Then
        LabelError.Caption = "Reminder Time occurs in the past."
        TextTime.SetFocus
        Exit Sub
    End If
    
    ButtonOK.Enabled = True
    
End Sub

Private Function ConvDate() As Variant

    If IsDate(TextDate.value) Then
        ConvDate = TextDate.value
        Exit Function
    End If
    
    '   SPOS - "a long date format is not recognized if it also contains the day-of-the-week string"
    '
    Dim sDate As String
    sDate = Trim(TextDate.value)

    Dim DaysShort() As Variant
    DaysShort = Array("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")
    Dim DaysLong() As Variant
    DaysLong = Array("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")
    
    Dim Day As Variant
    For Each Day In DaysShort
        sDate = Replace(sDate, Day, "")
    Next Day
    For Each Day In DaysLong
        sDate = Replace(sDate, Day, "")
    Next Day
    
    sDate = Trim(sDate)
    
    If IsDate(sDate) Then ConvDate = sDate

End Function

Private Function ConvTime() As Variant

    If IsDate(TextTime.value) Then
        ConvTime = TextTime.value
        Exit Function
    End If
    
End Function

Private Sub DatePicker_Click()

    TextDate.value = Format( _
        CalendarForm.GetDate( _
            SelectedDate:=ConvDate(), _
            DateFontSize:=9, _
            TodayButton:=True, _
            OkayButton:=False _
        ), FormatDate)
    Validate

End Sub

