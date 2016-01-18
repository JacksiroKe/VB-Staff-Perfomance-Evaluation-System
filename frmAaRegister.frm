VERSION 5.00
Begin VB.Form frmAaRegister 
   Caption         =   "Register Your Account NOW !!!"
   ClientHeight    =   5610
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5640
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
   ScaleHeight     =   5610
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLoginNow 
      Caption         =   "Login Now"
      Height          =   615
      Left            =   480
      TabIndex        =   6
      Top             =   4680
      Width           =   4575
   End
   Begin VB.CommandButton cmdRegisterNow 
      Caption         =   "Register Now"
      Height          =   615
      Left            =   480
      TabIndex        =   4
      Top             =   3240
      Width           =   4575
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Text            =   "Preferred Password"
      Top             =   2520
      Width           =   4575
   End
   Begin VB.TextBox txtFullName 
      Alignment       =   2  'Center
      Height          =   480
      Left            =   480
      TabIndex        =   2
      Text            =   "First and Last Name"
      Top             =   1080
      Width           =   4575
   End
   Begin VB.TextBox txtUsername 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Text            =   "Preferred Username"
      Top             =   1800
      Width           =   4575
   End
   Begin VB.Label lblOr 
      Alignment       =   2  'Center
      Caption         =   "OR"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Register Your Account"
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5655
   End
End
Attribute VB_Name = "frmAaRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset

Dim account As String

Private Sub cmdLoginNow_Click()
    frmAaLogin.Show
    Unload Me
End Sub

Private Sub cmdRegisterNow_Click()
    If Trim(txtPassword.Text) = "" Or Trim(txtPassword.Text) = "Preferred Password" Or Trim(txtUsername.Text) = "" Or Trim(txtUsername.Text) = "Preferred Username" Then
        MsgBox "Incomplete information provided, Enter all fields to continue", vbCritical, "Validation"
        Exit Sub
    End If
  On Error GoTo ErrorHandler
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from administrators", con, adOpenKeyset, adLockOptimistic
    Rs.AddNew
    Rs!username = txtUsername.Text
    Rs!fullname = txtFullName.Text
    Rs!password = txtPassword.Text
    Rs.Update
    Rs.Close
    MsgBox "You Have Registered succesfully as an Admin! Now Login", vbInformation, App.Title
    frmAaLogin.Show
    Unload Me
Exit Sub
ErrorHandler:
 MsgBox Err.Description & " No. " & Err.Number
End Sub

Private Sub Form_Load()
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\pucspes.mdb;"
    con.Open
End Sub

Private Sub txtFullName_Click()
    txtFullName.Text = ""
End Sub

Private Sub txtPassword_Click()
    txtPassword.Text = ""
End Sub

Private Sub txtUsername_Click()
    txtUsername.Text = ""
End Sub
