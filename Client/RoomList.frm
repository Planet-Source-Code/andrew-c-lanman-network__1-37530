VERSION 5.00
Begin VB.Form RoomList 
   BorderStyle     =   0  'None
   ClientHeight    =   4275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Create"
      Height          =   315
      Left            =   2715
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3870
      Width           =   1200
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   495
      Top             =   3240
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Refresh"
      Height          =   315
      Left            =   3915
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3870
      Width           =   1200
   End
   Begin VB.CommandButton CmdJoin 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Join"
      Height          =   315
      Left            =   1515
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3870
      Width           =   1200
   End
   Begin VB.ListBox RoomList 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
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
      Height          =   3390
      Left            =   120
      TabIndex        =   7
      Top             =   420
      Width           =   6270
   End
   Begin VB.PictureBox Picture4 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   6360
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   4020
      Width           =   195
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   4
      Top             =   4020
      Width           =   195
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   6345
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   255
      Width           =   195
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   255
      Width           =   195
   End
   Begin VB.PictureBox TopBarRight 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5835
      ScaleHeight     =   255
      ScaleWidth      =   705
      TabIndex        =   1
      Top             =   0
      Width           =   705
      Begin VB.Image Image9 
         Height          =   240
         Left            =   465
         Top             =   15
         Width           =   225
      End
   End
   Begin VB.PictureBox TopBarLeft 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   465
      TabIndex        =   0
      Top             =   0
      Width           =   465
   End
   Begin VB.Image Image8 
      Height          =   255
      Left            =   360
      Top             =   0
      Width           =   5655
   End
   Begin VB.Image Image6 
      Height          =   270
      Left            =   195
      Stretch         =   -1  'True
      Top             =   4020
      Width           =   6735
   End
   Begin VB.Image Image5 
      Height          =   3675
      Left            =   6345
      Stretch         =   -1  'True
      Top             =   525
      Width           =   195
   End
   Begin VB.Label TitleText 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Lanman Network - Chat Rooms"
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
      TabIndex        =   6
      Top             =   15
      Width           =   6435
   End
   Begin VB.Image Image2 
      Height          =   3675
      Left            =   0
      Stretch         =   -1  'True
      Top             =   525
      Width           =   195
   End
   Begin VB.Image Image1 
      Height          =   270
      Left            =   180
      Stretch         =   -1  'True
      Top             =   255
      Width           =   6705
   End
   Begin VB.Image TopBarTile 
      Height          =   255
      Left            =   465
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5925
   End
End
Attribute VB_Name = "RoomList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurrentList As Boolean
Public NumberofRooms As Long
Public RoomCounter As Long
Public TempCounter As Long

Public Sub CheckRoomList()
    If CurrentList = False Then
        SendServerCommand "GetRoomCount"
        CurrentList = True
    End If
End Sub

Private Sub Command1_Click()
    CurrentList = False
    RoomList.Clear
    SendServerCommand "GetRoomCount"
End Sub

Private Sub Command2_Click()
CreateRoom.Show
End Sub

Private Sub Image8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngReturnValue As Long
    If Button = 1 Then
        Call ReleaseCapture
        lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If
End Sub

Private Sub Image9_Click()
    Me.Hide
End Sub

Private Sub RoomList_DblClick()
    Call CmdJoin_Click
End Sub

Private Sub CmdJoin_Click()
    On Error Resume Next
    ' Search through and see if all rooms are used
    For X = 0 To 20
    If Checkers(X) = False Then GoTo CheckedOK
    Next X
    ErrorPopup "You have too many chat-rooms open." & vbCrLf & "Close a few before trying-again."
    Exit Sub
CheckedOK:
    SendServerCommand "JoinRoom" & Chr(1) & Right(RoomList.List(RoomList.ListIndex), Len(RoomList.List(RoomList.ListIndex)) - InStr(1, RoomList.List(RoomList.ListIndex), ")", vbTextCompare) - 1) & Chr(1) & "<null>"
End Sub

Private Sub Timer1_Timer()
    DoEvents
    If GetActiveWindow = MessageBoard.hwnd Then
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
