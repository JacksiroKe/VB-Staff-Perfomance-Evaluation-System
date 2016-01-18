VERSION 5.00
Begin VB.Form frmAaHome 
   Caption         =   "Welcome to PUC Staff Perfomance Evaluation System!"
   ClientHeight    =   6435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13695
   BeginProperty Font 
      Name            =   "News Gothic"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   13695
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Quick Menu"
      Height          =   4815
      Left            =   8280
      TabIndex        =   1
      Top             =   1200
      Width           =   4935
      Begin VB.CommandButton cmdLogOut 
         Caption         =   "Log Out"
         Height          =   735
         Left            =   480
         TabIndex        =   5
         Top             =   3840
         Width           =   3975
      End
      Begin VB.CommandButton cmdAddNew 
         Caption         =   "Add a New Staff"
         Height          =   800
         Left            =   480
         TabIndex        =   4
         Top             =   2760
         Width           =   3975
      End
      Begin VB.CommandButton cmdViewResults 
         Caption         =   "View Results"
         Height          =   800
         Left            =   360
         TabIndex        =   3
         Top             =   1680
         Width           =   4215
      End
      Begin VB.CommandButton cmdEvaluate 
         Caption         =   "Evaluate"
         Height          =   800
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   4215
      End
   End
   Begin VB.Image Image1 
      Height          =   4800
      Left            =   480
      Picture         =   "frmAaHome.frx":0000
      Top             =   1080
      Width           =   7200
   End
   Begin VB.Label Label1 
      Caption         =   "Welcome to PUC Staff Perfomance Evaluation System"
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   14175
   End
End
Attribute VB_Name = "frmAaHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddNew_Click()
    frmAaManage.Show
    Unload Me
End Sub

Private Sub cmdEvaluate_Click()
    frmAaEvaluate.Show
    Unload Me
End Sub

Private Sub cmdLogOut_Click()
    frmAaLogin.Show
    Unload Me
End Sub

Private Sub cmdViewResults_Click()
    frmAaResults.Show
    Unload Me
End Sub

Private Sub Form_Load()
    If frmAaLogin.cmbAccount.Text = "Student" Then
        cmdAddNew.Visible = False
        cmdModerate.Visible = False
    End If
End Sub
