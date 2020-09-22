VERSION 5.00
Begin VB.Form CreateRoom 
   BackColor       =   &H00575030&
   BorderStyle     =   0  'None
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00808080&
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2235
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   13
      Text            =   "<null>"
      Top             =   1620
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00575030&
      Caption         =   "Make this room private"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   165
      TabIndex        =   11
      Top             =   1290
      Width           =   3480
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00575030&
      Caption         =   "Make this room public"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   165
      TabIndex        =   10
      Top             =   915
      Value           =   -1  'True
      Width           =   3480
   End
   Begin VB.TextBox txtRoomname 
      BackColor       =   &H00808080&
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   1125
      TabIndex        =   9
      Top             =   465
      Width           =   3615
   End
   Begin VB.PictureBox TopBarLeft 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      Picture         =   "CreateRoom.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   465
      TabIndex        =   6
      Top             =   0
      Width           =   465
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   0
      Picture         =   "CreateRoom.frx":06A2
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
      Left            =   4740
      Picture         =   "CreateRoom.frx":09B4
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
      Picture         =   "CreateRoom.frx":0CC6
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   2745
      Width           =   195
   End
   Begin VB.PictureBox Picture4 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   4755
      Picture         =   "CreateRoom.frx":0FD8
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   2745
      Width           =   195
   End
   Begin VB.CommandButton CmdJoin 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Create"
      Height          =   315
      Left            =   1845
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2580
      Width           =   1200
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   210
      Top             =   2325
   End
   Begin VB.PictureBox TopBarRight 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4215
      Picture         =   "CreateRoom.frx":12EA
      ScaleHeight     =   255
      ScaleWidth      =   705
      TabIndex        =   5
      Top             =   0
      Width           =   705
      Begin VB.Image Image9 
         Height          =   240
         Left            =   465
         Top             =   15
         Width           =   225
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password (Optional):"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   705
      TabIndex        =   12
      Top             =   1650
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Room name:"
      ForeColor       =   &H00808080&
      Height          =   225
      Left            =   180
      TabIndex        =   8
      Top             =   495
      Width           =   1035
   End
   Begin VB.Label TitleText 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Lanman Network - Create Chat-Room"
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
      Width           =   4260
   End
   Begin VB.Image Image8 
      Height          =   255
      Left            =   360
      Top             =   0
      Width           =   4035
   End
   Begin VB.Image TopBarTile 
      Height          =   255
      Left            =   465
      Picture         =   "CreateRoom.frx":1CBC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3810
   End
   Begin VB.Image Image6 
      Height          =   270
      Left            =   165
      Picture         =   "CreateRoom.frx":1D42
      Stretch         =   -1  'True
      Top             =   2745
      Width           =   6735
   End
   Begin VB.Image Image5 
      Height          =   2730
      Left            =   4740
      Picture         =   "CreateRoom.frx":1DCC
      Stretch         =   -1  'True
      Top             =   525
      Width           =   195
   End
   Begin VB.Image Image2 
      Height          =   2685
      Left            =   0
      Picture         =   "CreateRoom.frx":1E36
      Stretch         =   -1  'True
      Top             =   525
      Width           =   195
   End
   Begin VB.Image Image1 
      Height          =   270
      Left            =   195
      Picture         =   "CreateRoom.frx":1EA0
      Stretch         =   -1  'True
      Top             =   255
      Width           =   6705
   End
End
Attribute VB_Name = "CreateRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdJoin_Click()
Dim RoomType As String
Dim Roomname As String
Dim PasswordString As String
If Option1.Value = True Then RoomType = "Public"
If Option2.Value = True Then RoomType = "Private"
Roomname = txtRoomname.Text
PasswordString = txtPassword.Text
SendServerCommand "CreateRoom" & Chr(1) & Roomname & Chr(1) & RoomType & Chr(1) & PasswordString
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
   txtPassword.Visible = False
   Label2.Visible = False
   txtPassword.Text = "<null>"
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
   txtPassword.Visible = True
   Label2.Visible = True
   If txtPassword.Text = "" Then txtPassword = "<null>"
End If
End Sub
