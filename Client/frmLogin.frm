VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server Login"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   4545
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Remember my username and password for the future."
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Default         =   -1  'True
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Password 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1200
      Width           =   3495
   End
   Begin VB.TextBox Username 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Text            =   "Guest"
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label Label4 
      Caption         =   "Don't have an account? Click here!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      MouseIcon       =   "frmLogin.frx":030A
      TabIndex        =   5
      Top             =   2280
      Width           =   4455
   End
   Begin VB.Label Label3 
      Caption         =   "Password:"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Username:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Please enter your username and password to connect to the Lanman network:"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public AwaitingLogin As Boolean
Private Sub Command1_Click()
On Error GoTo ErrorOccured
AwaitingLogin = True
MainForm.Server.Close
MainForm.Server.Connect ServerIP
Exit Sub
ErrorOccured:
MsgBox "Unable to connect to Lanman network" & vbCrLf & "The remote server cannot be found.", vbExclamation, "Connection failure"
Exit Sub
End Sub

Private Sub Form_Load()
Username.Text = GetKey("Lanman", "Settings", "Username", "")
Username.SelStart = Len(Username.Text)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
MainForm.Show
End Sub

Private Sub Label4_Click()
frmLogin.Hide
frmCreateLogin.Show
End Sub
