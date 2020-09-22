VERSION 5.00
Begin VB.Form ChatWnd 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Chat Window"
   ClientHeight    =   4830
   ClientLeft      =   1005
   ClientTop       =   2130
   ClientWidth     =   7095
   Icon            =   "Chat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   930
      Top             =   3450
   End
   Begin VB.ListBox UserList 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   4260
      Left            =   5190
      TabIndex        =   10
      Top             =   405
      Width           =   1770
   End
   Begin VB.TextBox Message 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   120
      MaxLength       =   1000
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   4350
      Width           =   4200
   End
   Begin VB.TextBox Conversation 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   3900
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   405
      Width           =   5025
   End
   Begin VB.PictureBox TopBarLeft 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   465
      TabIndex        =   7
      Top             =   0
      Width           =   465
   End
   Begin VB.PictureBox TopBarRight 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   6390
      ScaleHeight     =   255
      ScaleWidth      =   705
      TabIndex        =   6
      Top             =   0
      Width           =   705
      Begin VB.Image Image8 
         Height          =   210
         Left            =   450
         Top             =   30
         Width           =   210
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   255
      Width           =   195
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   6900
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   4
      Top             =   255
      Width           =   195
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   4560
      Width           =   195
   End
   Begin VB.PictureBox Picture4 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   6915
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   4560
      Width           =   195
   End
   Begin VB.CommandButton SendBtn 
      BackColor       =   &H00808080&
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   330
      Left            =   4365
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4350
      Width           =   795
   End
   Begin VB.Timer TimeTimer 
      Interval        =   1000
      Left            =   2220
      Top             =   3705
   End
   Begin VB.Image Image9 
      Height          =   255
      Left            =   240
      Top             =   0
      Width           =   6375
   End
   Begin VB.Image Image1 
      Height          =   270
      Left            =   195
      Stretch         =   -1  'True
      Top             =   255
      Width           =   6705
   End
   Begin VB.Image Image2 
      Height          =   4155
      Left            =   0
      Stretch         =   -1  'True
      Top             =   525
      Width           =   195
   End
   Begin VB.Label TitleText 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "Chat room - "
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
      Left            =   2775
      TabIndex        =   8
      Top             =   0
      Width           =   900
   End
   Begin VB.Image Image5 
      Height          =   4155
      Left            =   6900
      Stretch         =   -1  'True
      Top             =   525
      Width           =   195
   End
   Begin VB.Image Image6 
      Height          =   270
      Left            =   195
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   6735
   End
   Begin VB.Image TopBarTile 
      Height          =   255
      Left            =   465
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5925
   End
End
Attribute VB_Name = "ChatWnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
SendServerCommand "RoomRemove" & Chr(1) & Me.Tag
End Sub

Private Sub Image8_Click()
For X = 0 To 20
If RoomMatches(X) = Me.Tag Then
   Checkers(X) = False
End If
Next X
Unload Me
Me.Hide
End Sub

Private Sub Image9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long
If Button = 1 Then
   Call ReleaseCapture
   lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub

Private Sub Message_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   KeyCode = 0
   Call SendBtn_Click
End If
End Sub

Private Sub SendBtn_Click()
   SendServerCommand "RoomPost" & Chr(1) & Me.Tag & Chr(1) & frmLogin.Username.Text & ": " & Message.Text
   Message.Text = ""
   Message.SelStart = Len(Message.Text)
End Sub

Private Sub Timer1_Timer()
DoEvents
If GetActiveWindow = Me.hwnd Then
   TopBarLeft.Picture = SkinRes.TopBarLeft0.Picture
   TopBarTile.Picture = SkinRes.TopBarTile0.Picture
   TopBarRight.Picture = SkinRes.TopBarRight0.Picture
   TitleText.ForeColor = &HFFFFFF
   BlinkCountCheck = True
Else
   TopBarLeft.Picture = SkinRes.TopBarLeft1.Picture
   TopBarTile.Picture = SkinRes.TopbarTile1.Picture
   TopBarRight.Picture = SkinRes.TopbarRight1.Picture
   TitleText.ForeColor = &HC0C0C0
End If
End Sub

Private Sub TitleText_Change()
TitleText.Left = Me.Width - (Me.Width / 2) - TitleText.Width / 2
End Sub

Private Sub TopBarTile_Click()
Dim lngReturnValue As Long
If Button = 1 Then
   Call ReleaseCapture
   lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub
