VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MessageBoard 
   BorderStyle     =   0  'None
   Caption         =   "Message Board"
   ClientHeight    =   4335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6450
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "PopupMessage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6450
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView Messages 
      Height          =   3510
      Left            =   135
      TabIndex        =   12
      Top             =   390
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   6191
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   65535
      BackColor       =   8421504
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   10354
      EndProperty
   End
   Begin VB.Timer FocusTimer 
      Interval        =   1
      Left            =   255
      Top             =   3375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Open"
      Height          =   315
      Left            =   1425
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3930
      Width           =   1200
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Refresh"
      Height          =   315
      Left            =   5025
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3930
      Width           =   1200
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Delete"
      Height          =   315
      Left            =   3825
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3930
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Compose"
      Height          =   315
      Left            =   2625
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3930
      Width           =   1200
   End
   Begin VB.CommandButton SetFolder 
      BackColor       =   &H00C0C0C0&
      Caption         =   "View Private"
      Height          =   315
      Left            =   225
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3930
      Width           =   1200
   End
   Begin VB.PictureBox Picture4 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   6270
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   11
      Top             =   4080
      Width           =   195
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   10
      Top             =   4080
      Width           =   195
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   6255
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   255
      Width           =   195
   End
   Begin VB.PictureBox TopBarRight 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5745
      ScaleHeight     =   255
      ScaleWidth      =   705
      TabIndex        =   6
      Top             =   0
      Width           =   705
      Begin VB.Image Image9 
         Height          =   225
         Left            =   195
         Top             =   15
         Width           =   240
      End
      Begin VB.Image Image8 
         Height          =   210
         Left            =   450
         Top             =   30
         Width           =   240
      End
   End
   Begin VB.PictureBox TopBarLeft 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   465
      TabIndex        =   5
      Top             =   0
      Width           =   465
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3270
      Top             =   2835
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PopupMessage.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PopupMessage.frx":065E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PopupMessage.frx":09B2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image7 
      Height          =   255
      Left            =   375
      Top             =   0
      Width           =   5535
   End
   Begin VB.Image Image6 
      Height          =   270
      Left            =   195
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   6075
   End
   Begin VB.Image Image5 
      Height          =   3675
      Left            =   6255
      Stretch         =   -1  'True
      Top             =   525
      Width           =   195
   End
   Begin VB.Label TitleText 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Lanman Network - Message Board"
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
      Left            =   30
      TabIndex        =   9
      Top             =   0
      Width           =   6345
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
      Left            =   195
      Stretch         =   -1  'True
      Top             =   255
      Width           =   6060
   End
   Begin VB.Image TopBarTile 
      Height          =   255
      Left            =   465
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5280
   End
End
Attribute VB_Name = "MessageBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Folder As String
Public FormActive As Boolean
Private Sub Command1_Click()
frmCompose.Show
End Sub

Private Sub Command2_Click()
On Error GoTo endmenow:
If Messages.SelectedItem.Key = "" Then Exit Sub
SendServerCommand "RemoveMessage" & Chr(1) & Messages.SelectedItem.Key & Chr(1) & frmLogin.Username.Text & Chr(1) & Folder
endmenow:
End Sub

Private Sub Command4_Click()
If Messages.SelectedItem.Key = "" Then Exit Sub
SendServerCommand "GetMessage" & Chr(1) & Messages.SelectedItem.Key & Chr(1) & frmLogin.Username.Text & Chr(1) & Folder
End Sub

Private Sub FocusTimer_Timer()
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

Private Sub Form_Activate()
If FormActive = False Then
   Call SetFolder_Click
   FormActive = True
End If
End Sub

Private Sub Form_Load()
Folder = "Private"
MessageCount = 0
CounterBuffer = 0
End Sub


Private Sub Image7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long
If Button = 1 Then
   Call ReleaseCapture
   lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub

Private Sub Image8_Click()
Me.Hide
End Sub

Private Sub Image9_Click()
Me.WindowState = 1
End Sub

Private Sub Messages_DblClick()
SendServerCommand "GetMessage" & Chr(1) & Messages.SelectedItem.Key & Chr(1) & frmLogin.Username.Text & Chr(1) & Folder
End Sub

Public Sub SetFolder_Click()
If SetFolder.Caption = "View Public" Then
   SetFolder.Caption = "View Private"
   Folder = "Public"
Else
   SetFolder.Caption = "View Public"
   Folder = "Private"
End If
Messages.ListItems.Clear
SendServerCommand "Get" & Folder & "Headers" & Chr(1) & frmLogin.Username.Text
End Sub

Public Sub LoadIntoPrivate(rawdata As String, pIndex As Integer)
IndexHeaders(pIndex) = rawdata
End Sub

Public Function GetHeaderDetail(pIndex As Integer)
GetHeaderDetail = IndexHeaders(pIndex)
End Function

