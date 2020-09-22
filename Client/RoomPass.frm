VERSION 5.00
Begin VB.Form RoomPass 
   BorderStyle     =   0  'None
   Caption         =   "   If wsIndex <> ""-1"" Then"
   ClientHeight    =   2310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   ScaleHeight     =   2310
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer AccessDeniedTMR 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1875
      Top             =   915
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   270
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   1035
      Width           =   4305
   End
   Begin VB.PictureBox TopBarLeft 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      Picture         =   "RoomPass.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   465
      TabIndex        =   6
      Top             =   0
      Width           =   465
   End
   Begin VB.PictureBox TopBarRight 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4260
      Picture         =   "RoomPass.frx":06A2
      ScaleHeight     =   255
      ScaleWidth      =   705
      TabIndex        =   5
      Top             =   0
      Width           =   705
      Begin VB.Image Image8 
         Height          =   210
         Left            =   450
         Top             =   30
         Width           =   240
      End
      Begin VB.Image Image9 
         Height          =   225
         Left            =   195
         Top             =   15
         Width           =   240
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   0
      Picture         =   "RoomPass.frx":1074
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   4
      Top             =   255
      Width           =   195
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   4770
      Picture         =   "RoomPass.frx":1386
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   255
      Width           =   195
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   0
      Picture         =   "RoomPass.frx":1698
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   2040
      Width           =   195
   End
   Begin VB.PictureBox Picture4 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   4785
      Picture         =   "RoomPass.frx":19AA
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   2040
      Width           =   195
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   1845
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1860
      Width           =   1200
   End
   Begin VB.Timer FocusTimer 
      Interval        =   1
      Left            =   180
      Top             =   1740
   End
   Begin VB.Label AccessDenied 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Access Denied!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   270
      Left            =   1260
      TabIndex        =   10
      Top             =   1410
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.Label PasswordText 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Password for %roomname% is required:"
      ForeColor       =   &H00808080&
      Height          =   570
      Left            =   150
      TabIndex        =   8
      Top             =   450
      Width           =   4635
   End
   Begin VB.Image Image1 
      Height          =   270
      Left            =   195
      Picture         =   "RoomPass.frx":1CBC
      Stretch         =   -1  'True
      Top             =   255
      Width           =   4620
   End
   Begin VB.Image Image2 
      Height          =   2115
      Left            =   0
      Picture         =   "RoomPass.frx":1D46
      Stretch         =   -1  'True
      Top             =   525
      Width           =   195
   End
   Begin VB.Image Image3 
      Height          =   1770
      Left            =   165
      Picture         =   "RoomPass.frx":1DB0
      Stretch         =   -1  'True
      Top             =   405
      Width           =   4635
   End
   Begin VB.Label TitleText 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Lanman Network - Private Chat-Room"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   60
      TabIndex        =   7
      Top             =   15
      Width           =   4890
   End
   Begin VB.Image Image5 
      Height          =   2205
      Left            =   4770
      Picture         =   "RoomPass.frx":1DF6
      Stretch         =   -1  'True
      Top             =   525
      Width           =   195
   End
   Begin VB.Image Image6 
      Height          =   270
      Left            =   195
      Picture         =   "RoomPass.frx":1E60
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   4620
   End
   Begin VB.Image Image7 
      Height          =   255
      Left            =   360
      Top             =   0
      Width           =   3930
   End
   Begin VB.Image TopBarTile 
      Height          =   255
      Left            =   465
      Picture         =   "RoomPass.frx":1EEA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "RoomPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Counter As Integer
Public Flasher As Integer
Public Sub AccessDeniedTMR_Timer()
On Error Resume Next
Select Case Flasher
Case 0
    Flasher = 1
    AccessDenied.Visible = True
Case 1
    Flasher = 0
    AccessDenied.Visible = False
    txtPassword.Text = ""
    txtPassword.SetFocus
End Select
End Sub

Private Sub Command1_Click()
    SendServerCommand "JoinRoom" & Chr(1) & Me.Tag & Chr(1) & txtPassword.Text
    Counter = 1
    Flasher = 0
End Sub
