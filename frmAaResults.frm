VERSION 5.00
Begin VB.Form frmAaResults 
   Caption         =   "PUC Staff Perfomance Evaluation Results"
   ClientHeight    =   8070
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11430
   BeginProperty Font 
      Name            =   "News Gothic"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   11430
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGoBack 
      Caption         =   "Go Back"
      Height          =   495
      Left            =   8280
      TabIndex        =   17
      Top             =   6960
      Width           =   2295
   End
   Begin VB.CommandButton cmdLogOut 
      Caption         =   "Log Out"
      Height          =   495
      Left            =   5640
      TabIndex        =   16
      Top             =   6960
      Width           =   2295
   End
   Begin VB.ListBox lstStaffMembers 
      Height          =   5760
      Left            =   600
      TabIndex        =   1
      Top             =   1920
      Width           =   4455
   End
   Begin VB.Frame fraResults 
      Caption         =   "Results view: "
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   5400
      TabIndex        =   0
      Top             =   1800
      Width           =   5415
      Begin VB.ListBox lstComments 
         BeginProperty Font 
            Name            =   "News Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1980
         Left            =   240
         TabIndex        =   14
         Top             =   2640
         Width           =   4935
      End
      Begin VB.Shape Shape7 
         Height          =   2535
         Left            =   120
         Top             =   2160
         Width           =   5175
      End
      Begin VB.Label Label4 
         Caption         =   "Latest Comments"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   2280
         Width           =   4935
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ratings"
         BeginProperty Font 
            Name            =   "News Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   1700
         Width           =   975
      End
      Begin VB.Label lblTotal 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "Totals"
         Height          =   615
         Left            =   480
         TabIndex        =   12
         Top             =   1350
         Width           =   975
      End
      Begin VB.Label lblRate2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         Caption         =   "Rate2"
         BeginProperty Font 
            Name            =   "News Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3840
         TabIndex        =   11
         Top             =   1440
         Width           =   645
      End
      Begin VB.Label lblRate1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Rate1"
         BeginProperty Font 
            Name            =   "News Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4560
         TabIndex        =   10
         Top             =   1440
         Width           =   645
      End
      Begin VB.Label lblRate3 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         Caption         =   "Rate3"
         BeginProperty Font 
            Name            =   "News Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3120
         TabIndex        =   9
         Top             =   1440
         Width           =   645
      End
      Begin VB.Label lblRate4 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "Rate4"
         BeginProperty Font 
            Name            =   "News Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2400
         TabIndex        =   8
         Top             =   1440
         Width           =   645
      End
      Begin VB.Label lblRate5 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Rate5"
         BeginProperty Font 
            Name            =   "News Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1560
         TabIndex        =   7
         Top             =   1440
         Width           =   645
      End
      Begin VB.Shape Shape5 
         Height          =   735
         Left            =   120
         Top             =   1305
         Width           =   5175
      End
      Begin VB.Shape Shape4 
         Height          =   855
         Left            =   120
         Top             =   360
         Width           =   5175
      End
      Begin VB.Label lblDepartment 
         Caption         =   "Dept"
         BeginProperty Font 
            Name            =   "News Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   4815
      End
      Begin VB.Label lblGender 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "G"
         Height          =   375
         Left            =   4560
         TabIndex        =   5
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblFullName 
         Caption         =   "Full Name"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   4695
      End
   End
   Begin VB.Shape Shape6 
      Height          =   975
      Left            =   5400
      Top             =   6720
      Width           =   5415
   End
   Begin VB.Label lblInstructions 
      Alignment       =   2  'Center
      Caption         =   "Click on a Name to View Full Details of the Results"
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   5400
      TabIndex        =   13
      Top             =   2880
      Width           =   5415
   End
   Begin VB.Label Label2 
      Caption         =   "Latest Results in a Descending order. Click on a Name to view Results"
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   1320
      Width           =   10095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PUC Staff Perfomance Evaluation Results"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   360
      Width           =   10695
   End
   Begin VB.Shape Shape3 
      Height          =   6615
      Left            =   480
      Top             =   1200
      Width           =   10455
   End
   Begin VB.Shape Shape2 
      Height          =   735
      Left            =   480
      Top             =   360
      Width           =   10455
   End
   Begin VB.Shape Shape1 
      Height          =   7700
      Left            =   360
      Top             =   240
      Width           =   10695
   End
End
Attribute VB_Name = "frmAaResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim Stafno As String


Private Sub Load_Latest()
lstStaffMembers.Clear
On Error GoTo ErrorHandler
 Set Rs = New ADODB.Recordset
    Rs.Open "Select * from staff_members ORDER BY rates ASC", con, adOpenKeyset, adLockOptimistic
    Do Until Rs.EOF
        lstStaffMembers.AddItem Rs!staffname
        Rs.MoveNext
    Loop
    Exit Sub
ErrorHandler:
MsgBox Err.Description & " No. " & Err.Number
End Sub

Private Sub cmdGoBack_Click()
    frmAaHome.Show
    Unload Me
End Sub

Private Sub cmdLogOut_Click()
    frmAaLogin.Show
    Unload Me
End Sub

Private Sub Form_Load()
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\pucspes.mdb;"
    con.Open
    Load_Latest
    fraResults.Visible = False
End Sub

Private Sub lstStaffMembers_Click()
    lblInstructions.Visible = False
    fraResults.Visible = True
    On Error GoTo ErrorHandler
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from staff_members WHERE staffname = '" & lstStaffMembers.Text & "'", con, adOpenKeyset, adLockOptimistic
    lblFullName = Rs!staffname
    lblDepartment = "Department: " & Rs!department
    lblGender = Rs!sex
    lblTotal = Rs!rates
    lblRate5 = "Rate5 " & Rs!rate5
    lblRate4 = "Rate4 " & Rs!rate4
    lblRate3 = "Rate3 " & Rs!rate3
    lblRate2 = "Rate2 " & Rs!rate2
    lblRate1 = "Rate1 " & Rs!rate1
    Stafno = Rs!staffno
    Load_Comments
Exit Sub
ErrorHandler:
MsgBox Err.Description & " No. " & Err.Number
End Sub

Private Sub Load_Comments()
lstComments.Clear
On Error GoTo ErrorHandler
 Set Rs = New ADODB.Recordset
    Rs.Open "Select * from feedback WHERE staffno = '" & Stafno & "'", con, adOpenKeyset, adLockOptimistic
    Do Until Rs.EOF
        lstComments.AddItem "## " & Rs!comment
        Rs.MoveNext
    Loop
    Exit Sub
ErrorHandler:
MsgBox Err.Description & " No. " & Err.Number
End Sub

