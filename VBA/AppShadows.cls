VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppShadows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents AppInspectors As Outlook.Inspectors
Attribute AppInspectors.VB_VarHelpID = -1
Private WithEvents AppExplorers As Outlook.Explorers
Attribute AppExplorers.VB_VarHelpID = -1

Private InspShadowsCollection As VBA.Collection
Private ExplShadowsCollection As VBA.Collection
'
    
' =====================================================================
'   Shadow Manager
'
'       Derived from http://www.vboffice.net/en/developers/inspector-wrapper-receive-events-of-multiple-emails/
'       and http://www.codekabinett.com/rdumps.php?Lang=2&targetDoc=vba-handling-events-indefinite-sources
'       to handle multiple Inspectors/Explorers events.
'
' =====================================================================

'   Property Get - Inspector Shadow by Inspector or Inspector Shadow Key
'
Public Property Get InspShadows(ByVal Index As Variant) As InspShadow
    
    Set InspShadows = Nothing
    
    Select Case VarType(Index)
    Case vbString, vbLong
        Dim ShadowKey As String: ShadowKey = CStr(Index)
        If InInspCollection(ShadowKey) Then Set InspShadows = InspShadowsCollection.Item(ShadowKey)
    Case vbObject
        If Not TypeOf Index Is Outlook.Inspector Then Stop: Exit Function
        For Each InspShadows In InspShadowsCollection
            If InspShadows.Inspector Is Index Then Exit Property
        Next InspShadows
    End Select

End Property

'   Collections Add/Remove
'
Public Sub InspShadowsAdd(ByVal InspShadow As InspShadow, ByVal InspShadowKey As String):       InspShadowsCollection.Add InspShadow, InspShadowKey:      End Sub
Public Sub ExplShadowsAdd(ByVal ExplShadow As ExplShadow, ByVal ExplShadowKey As String):       ExplShadowsCollection.Add ExplShadow, ExplShadowKey:      End Sub
Public Sub InspShadowsRemove(ByVal InspShadowKey As String):                                    InspShadowsCollection.Remove InspShadowKey:               End Sub
Public Sub ExplShadowsRemove(ByVal ExplShadowKey As String):                                    ExplShadowsCollection.Remove ExplShadowKey:               End Sub

'   Init
'
Public Function Initialize() As Boolean
Initialize = False

    '   Create an InspShadow for any open Inspectors
    '
    Set InspShadowsCollection = New VBA.Collection
    Dim Inspector As Outlook.Inspector
    For Each Inspector In Application.Inspectors
        AppInspectors_NewInspector Inspector
    Next Inspector

    '   Create an ExplShadow for any open Explorers
    '
    Set ExplShadowsCollection = New VBA.Collection
    Dim Explorer As Outlook.Explorer
    For Each Explorer In Application.Explorers
        AppExplorers_NewExplorer Explorer
    Next Explorer

    '   Hook the Application.Inspectors/Explorers collections
    '
    Set AppInspectors = Application.Inspectors
    Set AppExplorers = Application.Explorers
    
Initialize = True
End Function

'   Called on New Inspector event or App Startup
'
Private Sub AppInspectors_NewInspector(ByVal Inspector As Outlook.Inspector)

    Static InspShadowKey As Long
    
    If Inspector Is Nothing Then Exit Sub
    
    '   Create a new InspShadow instance
    '
    Dim InspShadow As InspShadow
    Set InspShadow = New InspShadow
    If Not InspShadow.Initialize(Inspector, CStr(InspShadowKey)) Then Stop: Exit Sub
    InspShadowKey = InspShadowKey + 1
    
End Sub

'   Called on New Explorer event or App Startup
'
Private Sub AppExplorers_NewExplorer(ByVal Explorer As Outlook.Explorer)

    Static ExplShadowKey As Long

    If Explorer Is Nothing Then Exit Sub

    '   Create a new ExplShadow instance
    '
    Dim ExplShadow As ExplShadow
    Set ExplShadow = New ExplShadow
    If Not ExplShadow.Initialize(Explorer, CStr(ExplShadowKey)) Then Stop: Exit Sub
    ExplShadowKey = ExplShadowKey + 1

End Sub

'   Is a Key String in the Collection?
'
Private Function InInspCollection(ByVal ShadowKey As String) As Boolean

    On Error Resume Next
        InspShadowsCollection.Item ShadowKey
        InInspCollection = (Err.Number = glbError_None)
    On Error GoTo 0

End Function

