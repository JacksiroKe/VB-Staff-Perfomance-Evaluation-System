VERSION 5.00
Begin VB.Form frmAaLogin 
   Caption         =   "Login to PUC SPES"
   ClientHeight    =   4470
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4665
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   4665
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register"
      Height          =   735
      Left            =   2640
      TabIndex        =   5
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   3480
      Width           =   1455
   End
   Begin VB.ComboBox cmbAccount 
      Height          =   555
      ItemData        =   "frmAaLogin.frx":0000
      Left            =   240
      List            =   "frmAaLogin.frx":000A
      TabIndex        =   3
      Text            =   "Login as a:"
      Top             =   1080
      Width           =   4215
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Height          =   555
      Left            =   240
      TabIndex        =   2
      Text            =   "password"
      Top             =   2640
      Width           =   4215
   End
   Begin VB.TextBox txtUsername 
      Alignment       =   2  'Center
      Height          =   555
      Left            =   240
      TabIndex        =   1
      Text            =   "username"
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "OR"
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Login to your Account"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmAaLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim accountype As String
Dim username As String
Dim password As String
Dim login As Boolean

Private Sub cmbAccount_Change()
    accountype = cmbAccount.Text
End Sub

Private Sub isStudent()
username = txtUsername.Text
password = txtPassword.Text
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from students WHERE username = '" & username & "' AND password = '" & password & "'", con, adOpenKeyset, adLockOptimistic
    Rs.Close
    login = True
End Sub

Private Sub isAdmin()
username = txtUsername.Text
password = txtPassword.Text
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from administrators WHERE username = '" & username & "' AND password = '" & password & "'", con, adOpenKeyset, adLockOptimistic
    Rs.Close
    
End Sub

Private Sub cmdLoginx_Click()
    On Error GoTo ErrorHandler
    If accountype = "Student" Then
       isStudent
       If login = True Then
        MsgBox "Welcome back " & txtUsername.Text
       Else
        MsgBox "You should register!"
        End If
    End If
    
    If accountype = "Administrator" Then
        isAdmin
        If login = True Then
            MsgBox "Welcome back " & txtUsername.Text
       Else
        MsgBox "You should register!"
        End If
        
    End If
    Exit Sub
ErrorHandler:
MsgBox Err.Description & " No. " & Err.Number
frmAaHome.Show
Unload Me
End Sub

Private Sub cmdLogin_Click()
If cmbAccount.Text = "" Or cmbAccount.Text = "Login as a:" Then
    MsgBox "You must choose an acount first", vbCritical
End If
If cmbAccount.Text = "Student" Or cmbAccount.Text = "Administrator" Then
    frmAaHome.Show
    Unload Me
End If
End Sub

Private Sub cmdRegister_Click()
    frmAaRDialog.Show
End Sub

Private Sub txtPassword_Click()
    txtPassword.Text = ""
End Sub

Private Sub txtUsername_Click()
    txtUsername.Text = ""
End Sub
