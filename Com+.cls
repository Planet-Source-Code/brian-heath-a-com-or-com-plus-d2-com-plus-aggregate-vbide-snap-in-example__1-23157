VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Com"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function ShellExecute _
    Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long _
    ) As Long
    'Window constants for the API
    Const SW_SHOWMAXIMIZED = 3
    Const SW_SHOWMINIMIZED = 2
    Const SW_SHOWNORMAL = 1
    'Local variables
    Private mstrURL As String
    Private mlngMax As Long
    Private mlngMin As Long
    Private mlngNormal As Long
    
    Property Let URL(strURL As String)
    mstrURL = strURL
End Property


Property Let Maximized(bMax As Boolean)


    If bMax Then
        mlngMax = SW_SHOWMAXIMIZED
    End If
End Property


Property Let Minimized(bMin As Boolean)


    If bMin Then
        mlngMin = SW_SHOWMINIMIZED
    End If
End Property


Property Let Normal(bNormal As Boolean)


    If bNormal Then
        mlngNormal = SW_SHOWNORMAL
    End If
End Property


Public Sub OpenURL()
    
    Call ShellExecute(0&, vbNullString, mstrURL, _
    vbNullString, "C:\", mlngMax Or mlngMin Or mlngNormal)
End Sub


Public Sub Mail()
    
    Call ShellExecute(0&, vbNullString, "mailto:" & mstrURL, _
    vbNullString, "C:\", mlngMax Or mlngMin Or mlngNormal)
End Sub


