VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private WithEvents ProjectItems As Outlook.Items
Attribute ProjectItems.VB_VarHelpID = -1

' =====================================================================
'   ItemSend - Must be in ThisOutlookSession
' =====================================================================
'
Private Sub Application_ItemSend(ByVal Item As Object, ByRef Cancel As Boolean):    Cancel = Not Event_SendItem(Item):  End Sub

' =====================================================================
'   Misc Events
'
'   - Have to be in a Class, but I don't what to add a whole new,
'   - tiny class, just for these.
'
' =====================================================================

Public Function MiscEventsHook() As Boolean
MiscEventsHook = False

    '   ProjectItems - To catch Project Items Add.
    '
    Set ProjectItems = Folders_KnownPath(glbKnownPath_Projects).Items
    
MiscEventsHook = True
End Function
    
Private Sub ProjectItems_ItemAdd(ByVal Item As Object):     Event_AddItem Item, Folders_KnownPath(glbKnownPath_Projects):  End Sub
'

' =====================================================================
'   Application_Startup
' =====================================================================

Private Sub Application_Startup()

    '   If running an AltVBA from my HotRod Script - notify and done
    '
    If Misc_UsingAltVBA(ShowStartupMsg:=True) Then Exit Sub
    
    '   Do it All
    '
    If Not Event_InitApplication() Then Stop: Exit Sub

End Sub
