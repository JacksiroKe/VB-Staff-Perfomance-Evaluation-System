VERSION 5.00
Begin VB.Form frmAaRDialog 
   Caption         =   "You are registering as who?"
   ClientHeight    =   3315
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   ClipControls    =   0   'False
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
   ScaleHeight     =   3315
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAsAdmin 
      Caption         =   "Administrator"
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Width           =   3735
   End
   Begin VB.CommandButton cmdAsStude 
      Caption         =   "Student"
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "OR"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Choose your Account"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmAaRDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAsAdmin_Click()
    frmAaRegister.Show
    Unload Me
    Unload frmAaLogin
End Sub

Private Sub cmdAsStude_Click()
    frmAaRegister1.Show
    Unload Me
    Unload frmAaLogin
End Sub

