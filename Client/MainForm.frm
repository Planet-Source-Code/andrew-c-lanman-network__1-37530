VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MainForm 
   BorderStyle     =   0  'None
   Caption         =   "LanMan"
   ClientHeight    =   4845
   ClientLeft      =   3780
   ClientTop       =   2070
   ClientWidth     =   4155
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "MainForm.frx":030A
   ScaleHeight     =   4845
   ScaleWidth      =   4155
   Begin VB.Timer ConnectRetry 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   135
      Top             =   2760
   End
   Begin VB.PictureBox TopBarRight 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3450
      ScaleHeight     =   255
      ScaleWidth      =   705
      TabIndex        =   10
      Top             =   0
      Width           =   705
      Begin VB.Image MinHotspot 
         Height          =   210
         Left            =   225
         Top             =   30
         Width           =   210
      End
      Begin VB.Image CloseHotspot 
         Height          =   210
         Left            =   465
         Top             =   30
         Width           =   210
      End
   End
   Begin VB.PictureBox TopBarLeft 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   465
      TabIndex        =   9
      Top             =   0
      Width           =   465
   End
   Begin VB.Timer CheckButtons 
      Interval        =   1
      Left            =   165
      Top             =   675
   End
   Begin VB.PictureBox FServeBTN 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   2985
      Picture         =   "MainForm.frx":0614
      ScaleHeight     =   615
      ScaleWidth      =   600
      TabIndex        =   5
      Top             =   855
      Width           =   600
   End
   Begin VB.PictureBox BoardBTN 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   1770
      Picture         =   "MainForm.frx":198E
      ScaleHeight     =   615
      ScaleWidth      =   600
      TabIndex        =   4
      Top             =   855
      Width           =   600
   End
   Begin VB.PictureBox RoomsBTN 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   555
      Picture         =   "MainForm.frx":2D08
      ScaleHeight     =   615
      ScaleWidth      =   600
      TabIndex        =   3
      Top             =   855
      Width           =   600
   End
   Begin MSWinsockLib.Winsock Server 
      Left            =   120
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   6168
   End
   Begin VB.Timer ScrollingText 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   675
      Top             =   3825
   End
   Begin VB.PictureBox ButtonBar 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   75
      ScaleHeight     =   270
      ScaleWidth      =   3975
      TabIndex        =   1
      Top             =   1815
      Width           =   3975
      Begin VB.Image AboutBTN 
         Height          =   255
         Left            =   2940
         Top             =   15
         Width           =   960
      End
      Begin VB.Image UsersBTN 
         Height          =   255
         Left            =   1995
         Top             =   15
         Width           =   930
      End
      Begin VB.Image SystemBTN 
         Height          =   255
         Left            =   1035
         Top             =   15
         Width           =   945
      End
      Begin VB.Image GeneralBTN 
         Height          =   255
         Left            =   60
         Top             =   15
         Width           =   930
      End
   End
   Begin VB.PictureBox LoginFrame 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1965
      Left            =   2160
      ScaleHeight     =   1935
      ScaleWidth      =   4155
      TabIndex        =   0
      Top             =   2640
      Visible         =   0   'False
      Width           =   4185
      Begin VB.Image LoginBTN 
         Height          =   465
         Left            =   1605
         Top             =   750
         Width           =   960
      End
   End
   Begin VB.Timer FocusTimer 
      Interval        =   1
      Left            =   45
      Top             =   4320
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   45
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   39
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":4082
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":5396
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image MoveHotspot 
      Height          =   300
      Left            =   270
      Top             =   0
      Width           =   3390
   End
   Begin VB.Label TitleText 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Lanman Network - <Crano>"
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
      TabIndex        =   11
      Top             =   0
      Width           =   4050
   End
   Begin VB.Image TopbarTile 
      Height          =   255
      Left            =   465
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3000
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "File Server"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   2775
      TabIndex        =   8
      Top             =   1455
      Width           =   1005
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Messages"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   1620
      TabIndex        =   7
      Top             =   1455
      Width           =   885
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Chat Rooms"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Left            =   255
      TabIndex        =   6
      Top             =   1455
      Width           =   1200
   End
   Begin VB.Label TextScroll 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Lanman is offline"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   225
      Left            =   180
      TabIndex        =   2
      Top             =   405
      Width           =   3765
   End
   Begin VB.Image Image3 
      Height          =   1935
      Left            =   0
      Top             =   255
      Width           =   4155
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public recX As Integer
Public recY As Integer
Public frmXPOS As Integer
Public frmYPOS As Integer
Dim CursorBuffer As POINTAPI
Dim EvalFrmPos As Boolean
Dim Dragging As Boolean

Private Sub AboutBtn_Click()
About.Show
End Sub

Private Sub AboutBTN_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ButtonBar.Picture = SkinRes.Button4.Picture
End Sub

Private Sub AboutBTN_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ButtonBar.Picture = SkinRes.Button0.Picture
End Sub

Private Sub BoardBTN_Click()
MessageBoard.Show
End Sub

Private Sub BoardBTN_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
BoardBTN.Picture = ImageList1.ListImages(1).Picture
End Sub

Private Sub BoardBTN_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
BoardBTN.Picture = ImageList1.ListImages(2).Picture
End Sub

Private Sub CheckButtons_Timer()
Select Case GetHandle
Case RoomsBTN.hwnd
    Label1.ForeColor = &HFFFF&
    Label2.ForeColor = &H808080
    Label3.ForeColor = &H808080
Case BoardBTN.hwnd
    Label1.ForeColor = &H808080
    Label2.ForeColor = &HFFFF&
    Label3.ForeColor = &H808080
Case FServeBTN.hwnd
    Label1.ForeColor = &H808080
    Label2.ForeColor = &H808080
    Label3.ForeColor = &HFFFF&
Case Else
    Label1.ForeColor = &H808080
    Label2.ForeColor = &H808080
    Label3.ForeColor = &H808080
End Select
End Sub

Private Sub CloseHotspot_Click()
Form_QueryUnload 0, 0
End Sub

Private Sub ConfigBTN_Click()
AppSettings.Show
End Sub

Private Sub ConnectRetry_Timer()
Call SkinRes.mnuLogon_Click
ConnectRetry.Enabled = False
End Sub

Private Sub FocusTimer_Timer()
DoEvents
If GetActiveWindow = MainForm.hwnd Then
   TopBarLeft.Picture = SkinRes.TopBarLeft0.Picture
   TopBarTile.Picture = SkinRes.TopBarTile0.Picture
   TopBarRight.Picture = SkinRes.TopBarRight0.Picture
   TitleText.ForeColor = &HFFFFFF
   If BlinkCountCheck = False Then
      BlinkCount = 0
   End If
   BlinkCountCheck = True
Else
   TopBarLeft.Picture = SkinRes.TopBarLeft1.Picture
   TopBarTile.Picture = SkinRes.TopbarTile1.Picture
   TopBarRight.Picture = SkinRes.TopbarRight1.Picture
   BlinkCountCheck = False
   TitleText.ForeColor = &HC0C0C0
End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
   About.Show
End If
End Sub
'
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case X
        Case 7725:
            MainForm.Show
            MainForm.WindowState = 0
            MainForm.ZOrder (0)
            FlashWindow MainForm.hwnd, 0
    End Select
If Button = vbLeftButton Then
    Call DragForm(Me)
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Debug.Print "Program Terminated"
On Local Error Resume Next
SaveKey "Lanman", "Settings", "XPos", MainForm.Left
SaveKey "Lanman", "Settings", "YPos", MainForm.Top
'Call DeleteRgn(lRegion)
RemoveIcon Me
TerminateForms
End
End Sub

Private Sub Image4_Click()
PopupMessage.Show
End Sub

Private Sub FServeBTN_Click()
frmFServe.Show
End Sub

Private Sub FServeBTN_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FServeBTN.Picture = ImageList1.ListImages(1).Picture
End Sub

Private Sub FServeBTN_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
FServeBTN.Picture = ImageList1.ListImages(2).Picture
End Sub

Private Sub GeneralBTN_Click()
PopupMenu SkinRes.mnuGeneral, 0, GeneralBTN.Left + 90, 2090
End Sub

Private Sub GeneralBTN_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ButtonBar.Picture = SkinRes.Button1.Picture
End Sub

Private Sub GeneralBTN_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ButtonBar.Picture = SkinRes.Button0.Picture
End Sub

Private Sub LoginBTN_Click()
If Resource.onlinestatus.Caption = "false" Then
    frmLogin.Password.Text = ""
    frmLogin.Show
    MainForm.Hide
    frmLogin.Password.SetFocus
    'LogOn
Else
    LogOff
End If
End Sub

Private Sub LogoutBTN_Click()
If Resource.onlinestatus.Caption = "false" Then
    LogOn
Else
    LogOff
End If
End Sub

Private Sub MinHotspot_Click()
MainForm.WindowState = 1
MainForm.Hide
End Sub

Private Sub MoveHotspot_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
GetCursorPos CursorBuffer
    Select Case X
        Case 7725:
            MainForm.Show
            MainForm.WindowState = 0
            MainForm.ZOrder (0)
    End Select
Dim lngReturnValue As Long
If Button = 1 Then
   Call ReleaseCapture
   lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub

Private Sub MoveHotspot_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dragging = False
End Sub

Private Sub RoomsBTN_Click()
RoomList.Show
RoomList.CheckRoomList
End Sub

Private Sub RoomsBTN_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RoomsBTN.Picture = ImageList1.ListImages(1).Picture
End Sub

Private Sub RoomsBTN_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RoomsBTN.Picture = ImageList1.ListImages(2).Picture
End Sub

Private Sub ScrollingText_Timer()
If BlinkCountCheck = True Then
   If BlinkCount = 4 Then
      ScrollingText.Enabled = False
      TextScroll.Caption = ""
      TextScroll.Tag = ""
      Exit Sub
   End If
End If
If TextScroll.Tag = "1" Then
   TextScroll.Caption = TextString
   TextScroll.Tag = "0"
   BlinkCount = BlinkCount + 1
Else
   TextScroll.Caption = ""
   TextScroll.Tag = "1"
   BlinkCount = BlinkCount + 1
End If
End Sub

Private Sub Server_Close()
Server.Close
LogOff
End Sub

Private Sub Server_Connect()
If frmCreateLogin.CreateLogin = True Then
    SendServerCommand "CreateLogin" & Chr(1) & frmCreateLogin.Username.Text & Chr(1) & frmCreateLogin.Password.Text
Else
    SendServerCommand "RequestLogin"
End If
End Sub

Private Sub Server_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
Server.GetData strData
ProcessCommand strData
End Sub

Private Sub Server_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Select Case Number
Case 10061
If frmLogin.AwaitingLogin = True Then
   'MsgBox "Unable to connect to the Lanman network server!" & vbCrLf & "Check your internet connection and try again.", vbCritical, "Connection failure"
   Server.Close
   ConnectRetry.Enabled = True
End If
End Select
End Sub

Private Sub SystemBTN_Click()
PopupMenu SkinRes.mnuSystem, 0, SystemBTN.Left + 90, 2090
End Sub

Private Sub SystemBTN_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ButtonBar.Picture = SkinRes.Button2.Picture
End Sub

Private Sub SystemBTN_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ButtonBar.Picture = SkinRes.Button0.Picture
End Sub

Private Sub TextScroll_Click()
ScrollingText.Enabled = False
Spaces = 0
Packing = False
PackPOS = 0
TextPOS = 1
'MainForm.TextScroll.Alignment = 1
MainForm.TextScroll.Caption = ""
End Sub

Public Sub TitleBar_Click()
'// This is placeholder text
End Sub

Private Sub TitleBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
End Sub

Private Sub UserList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu SkinRes.mnuGeneral, , UserList.SelectedItem.Left + 150, UserList.SelectedItem.Top + 900
End If
End Sub

Private Sub UserList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PopupMessage.ZOrder (0)
End Sub

Private Sub UsersBTN_Click()
PopupMenu SkinRes.mnuUsers, 0, UsersBTN.Left + 90, 2090
End Sub

Private Sub UsersBTN_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ButtonBar.Picture = SkinRes.Button3.Picture
End Sub

Private Sub UsersBTN_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ButtonBar.Picture = SkinRes.Button0.Picture
End Sub
