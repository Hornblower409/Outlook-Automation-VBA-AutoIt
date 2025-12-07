Attribute VB_Name = "Ribbon_"
Option Explicit
Option Private Module

'   Can a Window Ribbon enable an idMSO command?
'
Public Function Ribbon_EnabledMSO(ByVal Window As Object, ByVal idMSO As String) As Boolean
Ribbon_EnabledMSO = False

    '   Window must be an Inspector or Explorer
    '
    If Window Is Nothing Then Exit Function
    If Not ((TypeOf Window Is Outlook.Inspector) Or (TypeOf Window Is Outlook.Explorer)) Then Exit Function
    
    '   Get the Window CommandBars
    '
    Dim Commandbars As Office.Commandbars
    Set Commandbars = Window.Commandbars
    
    '   If this window can not enable the command - done
    '
    Dim EnabledMSO As Boolean:  EnabledMSO = False
    On Error Resume Next
        EnabledMSO = Commandbars.GetEnabledMso(idMSO)
        Select Case Err.Number
            Case glbError_None, glbError_InvalidProcArg, glbError_ObjectNoActionSupport
            Case Else
                Stop: Exit Function
        End Select
    On Error GoTo 0
    If Not EnabledMSO Then Exit Function
    
Ribbon_EnabledMSO = True
End Function

'   Is an Inspector Window Ribbon Active?
'
Public Function Ribbon_Active(ByVal Window As Object) As Boolean

    '   2024-11-17 - Explorers will not have File -> Save enabled unless
    '   there is an item selected in the view. So I'm trying File -> Options instead.
    '
    If TypeOf Window Is Outlook.Explorer Then
        Ribbon_Active = Ribbon_EnabledMSO(Window, glbidMSO_ApplicationOptionsDialog)
    ElseIf TypeOf Window Is Outlook.Inspector Then
        Ribbon_Active = Ribbon_EnabledMSO(Window, glbidMSO_FileSaveAs)
    Else
        '   You need to look at this one
        '
        Ribbon_Active = False
        Stop: Exit Function
    End If
    
End Function

'   Execute an idMSO command on a Window Ribbon
'
Public Function Ribbon_ExecuteMSO(ByVal Window As Object, ByVal idMSO As String) As Boolean
Ribbon_ExecuteMSO = False

    If Not Ribbon_Active(Window) Then Exit Function
    If Not Ribbon_EnabledMSO(Window, idMSO) Then Exit Function
    Window.Commandbars.ExecuteMso idMSO

Ribbon_ExecuteMSO = True
End Function

'   Is an Inspector in Edit Mode?
'
Public Function Ribbon_EditModeActive(ByVal Inspector As Outlook.Inspector) As Boolean
Ribbon_EditModeActive = False

    '   By trial-and-error:
    '
    '   The Window must be active
    '   Insert -> AttachFile is Enabled (the Message is enabled for editing)
    '   Actions -> Edit Message is Disabled (the Menu disables once you click it)
    '   Then the Item is in Edit Mode
    '
    '   You can't test just for Edit Message is Disabled. It's disabled in lots of other cases.
    '
    If Not Ribbon_EnabledMSO(Inspector, glbidMSO_AttachFile) Then Exit Function
    If Ribbon_EnabledMSO(Inspector, glbidMSO_EditMessage) Then Exit Function

Ribbon_EditModeActive = True
End Function

'   Put an Inspector in Edit Mode. TRUE <- Success
'
Public Function Ribbon_EditModeActivate(ByVal oInspector As Outlook.Inspector) As Boolean
Ribbon_EditModeActivate = False

    '   If Inspector.Open has been Canceled - Done
    '   If it is already in Edit Mode - Done
    '
    If oInspector.CurrentItem Is Nothing Then Ribbon_EditModeActivate = True: Exit Function
    If Ribbon_EditModeActive(oInspector) Then Ribbon_EditModeActivate = True: Exit Function
    
    '   Try to put it in Edit Mode
    '   Test if it worked
    '
    If Not Ribbon_ExecuteMSO(oInspector, glbidMSO_EditMessage) Then Exit Function
    If Not Ribbon_EditModeActive(oInspector) Then Exit Function
    
Ribbon_EditModeActivate = True
End Function

'   Is an Inspector in Design Mode?
'
'   !!  The Inspector MUST be out of the Open Event Scope  !!
'
Public Function Ribbon_DesignModeActive(ByVal oInspector As Outlook.Inspector) As Boolean
Ribbon_DesignModeActive = False

    If Not Ribbon_EnabledMSO(oInspector, glbidMSO_ViewVisualBasicCode) Then Exit Function

Ribbon_DesignModeActive = True
End Function

' ---------------------------------------------------------------------
'   Workaround for my Custom Forms not having a "Reply All".
' ---------------------------------------------------------------------
'
Public Function Ribbon_ReplyAll() As Boolean
Const ThisProc = "Ribbon_ReplyAll"
Ribbon_ReplyAll = True

    '   If Window has a Reply All - Do it and done.
    '   Else if Window has just Reply - Do it and done
    '
    If Ribbon_ExecuteMSO(ActiveWindow, glbidMSO_ReplyAll) Then Exit Function
    If Ribbon_ExecuteMSO(ActiveWindow, glbidMSO_Reply) Then Exit Function
   
Ribbon_ReplyAll = False
End Function

' ---------------------------------------------------------------------
'   Execute an Control via Control.Execute
'
'       This is the Old School way of executing a menu item. Should use Ribbon_ExecuteMSO
'       if at all possible. The Web says this is going away in newer Office.
'
' ---------------------------------------------------------------------
'
Public Function Ribbon_ControlExecute(ByVal Window As Object, ByVal ControlID As Long) As Boolean
Ribbon_ControlExecute = False

    Dim Commandbars As Office.Commandbars
    Set Commandbars = Window.Commandbars

    Dim CommandBar As Office.CommandBar
    For Each CommandBar In Commandbars

        Dim Control As CommandBarControl
        Set Control = CommandBar.FindControl(ID:=ControlID, Recursive:=True)
        If Not Control Is Nothing Then Exit For

    Next CommandBar
    
    '   If the Control was not found or is not visable - bail out
    '
    If Control Is Nothing Then Exit Function
    If Not Control.Visible Then Exit Function
    
    '   Click it
    '
    On Error Resume Next
        Control.Execute
        Select Case Err.Number
            Case glbError_None
            Case glbError_OpenRecurringFromInstance
                Exit Function
            Case Else
                Stop: Exit Function
        End Select
    On Error GoTo 0

Ribbon_ControlExecute = True
End Function

