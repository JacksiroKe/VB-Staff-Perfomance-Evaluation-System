VERSION 5.00
Begin VB.Form frmAaSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7905
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrTimer1 
      Interval        =   5000
      Left            =   360
      Top             =   4080
   End
   Begin VB.Label Label3 
      Caption         =   "STUDENT NAME: 	WINFRIDAH MORAGWA MAUTI	"
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   2760
      Width           =   6495
   End
   Begin VB.Label Label4 
      Caption         =   "ADMISSION NUMBER: 2014/CS/29535	"
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   3240
      Width           =   6135
   End
   Begin VB.Label Label8 
      Caption         =   "SUPERVISOR: MR.OKIDIA"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   5160
      Width           =   6375
   End
   Begin VB.Label Label7 
      Caption         =   "COURSE CODE:"
      Height          =   855
      Left            =   960
      TabIndex        =   4
      Top             =   4680
      Width           =   6255
   End
   Begin VB.Label Label6 
      Caption         =   "DEPARTMENT: COMPUTER STUDIES	"
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   4200
      Width           =   6255
   End
   Begin VB.Label Label5 
      Caption         =   "PROGRAMME NAME: COMPUTER STUDIES	"
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   3720
      Width           =   6255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Based on Student's Feedback"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      TabIndex        =   1
      Top             =   1920
      Width           =   5895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "PUC STAFF PERFORMANCE EVALUATION SYSTEM "
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   6015
   End
End
Attribute VB_Name = "frmAaSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub tmrTimer1_Timer()
    tmrTimer1.Enabled = False
    frmAaLogin.Show
    Unload Me
End Sub
