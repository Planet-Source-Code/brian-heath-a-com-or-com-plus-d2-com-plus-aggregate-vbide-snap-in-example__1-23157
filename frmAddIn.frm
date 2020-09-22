VERSION 5.00
Begin VB.Form frmAddIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "d2 Com Plus Aggregate "
   ClientHeight    =   3435
   ClientLeft      =   1845
   ClientTop       =   2580
   ClientWidth     =   6030
   Icon            =   "frmAddIn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHide 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton SUButton 
      Caption         =   "Start &Up"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2205
      ItemData        =   "frmAddIn.frx":038A
      Left            =   360
      List            =   "frmAddIn.frx":038C
      TabIndex        =   1
      Top             =   480
      Width           =   3855
   End
   Begin VB.CommandButton SDButton 
      Caption         =   "Shut &Down"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   4680
      Picture         =   "frmAddIn.frx":038E
      Top             =   2040
      Width           =   1125
   End
   Begin VB.Label URL 
      Alignment       =   2  'Center
      Caption         =   "http://www.deffacto.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2880
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "COM+ Application Controller"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VBInstance As VBIDE.VBE
Public Connect    As Connect
Dim NTCat         As Catalog
Dim NTApps        As CatalogCollection
Dim NTItem        As CatalogObject
Dim MTSCat        As COMAdminCatalog
Dim Apps          As COMAdminCatalogCollection
Dim item          As COMAdminCatalogObject
Dim version       As String
Option Explicit

Private Sub cmdHide_Click()
     Unload Me
End Sub

Private Sub Form_Load()
   
    On Error GoTo nt_code
   
   Set MTSCat = New COMAdminCatalog
   
   version = "w2k"
   
   Set Apps = MTSCat.GetCollection("Applications")
   
  Apps.Populate
   
   For Each item In Apps
      List1.AddItem item.Name
   Next
   
   Exit Sub
   
nt_code:
      version = "nt"
      
      Set NTCat = New Catalog
      
      Set NTApps = NTCat.GetCollection("Packages")
      
      NTApps.Populate
      
      For Each NTItem In NTApps
         List1.AddItem NTItem.Name
      Next
   
End Sub

Private Sub SDButton_Click()
  
   On Error GoTo errorhandler
   Dim utilItem As PackageUtil
   Dim pkid As String
   
   If version = "nt" Then
      Set utilItem = NTApps.GetUtilInterface
      
      For Each NTItem In NTApps
         If NTItem.Name = List1.Text Then
            pkid = NTItem.Value("ID")
         End If
      Next
      utilItem.ShutdownPackage pkid
   Else
   
    MTSCat.ShutdownApplication (List1.Text)
    MsgBox "Application Terminated Correctly"
   End If
   
   Exit Sub
   
errorhandler:
   MsgBox (Err.Description)
End Sub

Private Sub SUButton_Click()

  On Error GoTo errorhandler
  Dim utilItem As PackageUtil
  Dim pkid As String
   
   If version = "nt" Then
      Set utilItem = NTApps.GetUtilInterface
      
      For Each NTItem In NTApps
         If NTItem.Name = List1.Text Then
            pkid = NTItem.Value("ID")
         End If
      Next
      'utilItem.ShutdownPackage pkid
      MsgBox "NT STARTUP OPTION NOT FUNCTIONAL NOW, USE WIN2K"
   Else
   
   MTSCat.StartApplication (List1.Text)
    MsgBox "Start Up Of Application Was Successful"
   End If
   Exit Sub
   
errorhandler:
   MsgBox (Err.Description)
End Sub


Private Sub URL_Click()

On Error GoTo errorhandler
Dim Hlink As New Com

     Hlink.URL = "http://www.deffacto.com"
    'write the email address if you want to send an email
     Hlink.Maximized = True
     Hlink.OpenURL
     
     Exit Sub
     
errorhandler:
   MsgBox (Err.Description)
End Sub
