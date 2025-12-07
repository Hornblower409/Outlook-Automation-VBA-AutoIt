Attribute VB_Name = "TEST_"
Option Explicit

' NOT Option Private Module - I want these Subs to be exposed

Public glbTestClass As ANG.TEST_Class

' =====================================================================
'   2025-11-21 - Modified from source at
'   https://superuser.com/questions/1168978/how-can-i-automate-a-rule-that-runs-to-run-all-rules-on-inbox-after-an-imap-m
' =====================================================================

Sub RunAllInboxRules()

    Dim st As Outlook.Store
    Dim myRules As Outlook.Rules
    Dim rl As Outlook.Rule
    Dim count As Integer
    Dim ruleList As String

    ' Walk all Stores
    Dim oStores As Outlook.Stores
    Set oStores = Application.Session.Stores
    For Each st In oStores: Do

        On Error Resume Next
            Set myRules = st.GetRules
            Select Case Err.Number
                ' No Error - Continue
                Case 0
                ' Store has no Rules - Next Store
                Case -2147352567
                    Exit Do
                ' Else - Boom
                Case Else
                    Stop: Exit Sub
            End Select
        On Error GoTo 0
        
        ' Walk all Rules
        For Each rl In myRules: Do
        
            ' Must be Enabled
            If Not rl.Enabled Then Exit Do
            
            ' Must be Incoming
            If Not rl.RuleType = olRuleReceive Then Exit Do
            
            ' Must be Local i.e. "On this computer only"
            If Not rl.IsLocalRule Then Exit Do
            
            ' Run the Rule
            rl.Execute ShowProgress:=True
            count = count + 1
            ruleList = ruleList & vbCrLf & rl.Name
            
        Loop While False: Next rl
    
    Loop While False: Next st

    ' tell the user what you did
    ruleList = "These Local rules were executed against the Inbox: " & vbCrLf & ruleList
    MsgBox ruleList, vbInformation, "Macro: RunAllInboxRules"

End Sub
