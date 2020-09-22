VERSION 5.00
Begin VB.Form frmCreateLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Login Account"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   Icon            =   "frmCreateLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   5025
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Agreement 
      Height          =   1605
      Left            =   390
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "frmCreateLogin.frx":030A
      Top             =   1680
      Width           =   4140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go!"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox Password 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1200
      Width           =   3975
   End
   Begin VB.TextBox Username 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Text            =   "Guest"
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   5535
   End
   Begin VB.Label Label2 
      Caption         =   "Before an account can be created, you must agree to the terms and conditions of use."
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label Label1 
      Caption         =   "Username:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   5535
   End
End
Attribute VB_Name = "frmCreateLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CreateLogin As Boolean
Private Sub Command1_Click()
If MainForm.Server.State <> 7 Then
   MainForm.Server.Close
   CreateLogin = True
   MainForm.Server.Connect ServerIP
Else
   SendServerCommand "CreateLogin" & Chr(1) & Username.Text & Chr(1) & Password.Text
End If
End Sub

Private Sub Form_Load()
Randomize
Username.Text = "Guest" & Int(Rnd * 65535)
End Sub
