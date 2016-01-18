VERSION 5.00
Begin VB.Form frmAaManage 
   Caption         =   "Administrator - Manage a Staff Member"
   ClientHeight    =   6570
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14775
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
   ScaleHeight     =   6570
   ScaleWidth      =   14775
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLogout 
      Caption         =   "Log Out"
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   13
      Top             =   5880
      Width           =   2535
   End
   Begin VB.CommandButton cmdGoBack 
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   12
      Top             =   5880
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      Caption         =   "Staff Members: "
      Height          =   6135
      Left            =   6720
      TabIndex        =   9
      Top             =   240
      Width           =   7575
      Begin VB.ListBox lstStaffMembers 
         Height          =   4620
         Left            =   480
         TabIndex        =   11
         Top             =   1200
         Width           =   6615
      End
      Begin VB.ComboBox cmbDept2 
         Height          =   405
         Left            =   480
         TabIndex        =   10
         Text            =   "Sort by Department"
         Top             =   480
         Width           =   6615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add a New Staff Member: "
      Height          =   4815
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   5415
      Begin VB.CommandButton cmdAddNew 
         Caption         =   "Add this Staff Member"
         Height          =   615
         Left            =   600
         TabIndex        =   8
         Top             =   3720
         Width           =   4335
      End
      Begin VB.OptionButton optSex 
         Caption         =   "Female"
         Height          =   495
         Index           =   1
         Left            =   3120
         TabIndex        =   6
         Top             =   3120
         Width           =   1335
      End
      Begin VB.OptionButton optSex 
         Caption         =   "Male"
         Height          =   495
         Index           =   0
         Left            =   1560
         TabIndex        =   5
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox txtStaffNo 
         Height          =   495
         Left            =   480
         TabIndex        =   4
         Text            =   "Staff No"
         Top             =   2280
         Width           =   4575
      End
      Begin VB.TextBox txtFullName 
         Height          =   495
         Left            =   480
         TabIndex        =   3
         Text            =   "First and Last Name"
         Top             =   1320
         Width           =   4575
      End
      Begin VB.ComboBox cmbDept1 
         Height          =   405
         ItemData        =   "frmAaManage.frx":0000
         Left            =   480
         List            =   "frmAaManage.frx":001C
         TabIndex        =   2
         Text            =   "Choose a Department"
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Label2 
         Caption         =   "Sex:"
         Height          =   495
         Left            =   480
         TabIndex        =   7
         Top             =   3120
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Administrator - Add Staff Members"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   9855
   End
End
Attribute VB_Name = "frmAaManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset

Dim sex As String
    
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
End Sub

Private Sub cmdAddNew_Click()
    
    If Trim(txtFullName.Text) = "" Or Trim(txtFullName.Text) = "First and Last Name" Or Trim(txtStaffNo.Text) = "" Or Trim(txtStaffNo.Text) = "Staff No" Or cmbDept1.Text = "Choose a Department" Then
        MsgBox "Incomplete information provided, Enter all fields to continue", vbCritical, "Validation"
        Exit Sub
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from staff_members", con, adOpenKeyset, adLockOptimistic
    Rs.AddNew
    Rs!staffname = txtFullName.Text
    Rs!staffno = txtStaffNo.Text
    Rs!department = cmbDept1.Text
    Rs!sex = sex
    Rs.Update
    Rs.Close
    MsgBox "A new member of staff was added succesfully", vbInformation, App.Title
    Load_Latest
    
End Sub

Private Sub Load_Latest()
lstStaffMembers.Clear
On Error GoTo ErrorHandler
 Set Rs = New ADODB.Recordset
    Rs.Open "Select * from staff_members ORDER BY staffid ASC", con, adOpenKeyset, adLockOptimistic
    Do Until Rs.EOF
        lstStaffMembers.AddItem Rs!staffname
        Rs.MoveNext
    Loop
    Exit Sub
ErrorHandler:
MsgBox Err.Description & " No. " & Err.Number
End Sub

Private Sub optSex_Click(Index As Integer)
    If (Index = 1) Then
        sex = "female"
    Else
        sex = "male"
    End If
End Sub

Private Sub txtFullName_Click()
    txtFullName.Text = ""
End Sub

Private Sub txtStaffNo_Click()
    txtStaffNo.Text = ""
End Sub
