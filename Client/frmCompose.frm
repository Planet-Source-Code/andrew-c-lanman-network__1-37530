VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCompose 
   BorderStyle     =   0  'None
   Caption         =   "Compose New"
   ClientHeight    =   6195
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6195
   Icon            =   "frmCompose.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2880
      Top             =   5415
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00504C33&
      BorderStyle     =   0  'None
      Height          =   3645
      Left            =   105
      TabIndex        =   7
      Top             =   1410
      Width           =   5970
      Begin VB.TextBox MessageBody 
         BackColor       =   &H00808080&
         ForeColor       =   &H00C0C0C0&
         Height          =   3375
         Left            =   75
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   180
         Width           =   5805
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00504C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   5970
      Begin VB.TextBox Subject 
         BackColor       =   &H00808080&
         ForeColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   705
         TabIndex        =   6
         Top             =   705
         Width           =   5085
      End
      Begin VB.CheckBox PublicOption 
         BackColor       =   &H00504C33&
         Caption         =   "Place this message on the public message board"
         ForeColor       =   &H00C0C0C0&
         Height          =   240
         Left            =   705
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   10
         Top             =   345
         Width           =   4290
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Emails"
         Height          =   285
         Left            =   5055
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   330
         Width           =   735
      End
      Begin VB.TextBox txtTo 
         BackColor       =   &H00808080&
         ForeColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   705
         TabIndex        =   3
         Top             =   15
         Width           =   4305
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Users"
         Height          =   285
         Left            =   5055
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   15
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject:"
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   30
         TabIndex        =   5
         Top             =   720
         Width           =   4710
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "To:"
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   45
         Width           =   2415
      End
   End
   Begin VB.PictureBox TopBarLeft 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      Picture         =   "frmCompose.frx":030A
      ScaleHeight     =   255
      ScaleWidth      =   465
      TabIndex        =   21
      Top             =   0
      Width           =   465
   End
   Begin VB.PictureBox TopBarRight 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5505
      Picture         =   "frmCompose.frx":09AC
      ScaleHeight     =   255
      ScaleWidth      =   705
      TabIndex        =   20
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
      Picture         =   "frmCompose.frx":137E
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   19
      Top             =   255
      Width           =   195
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   6015
      Picture         =   "frmCompose.frx":1690
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   18
      Top             =   255
      Width           =   195
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   -15
      Picture         =   "frmCompose.frx":19A2
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   17
      Top             =   5925
      Width           =   195
   End
   Begin VB.PictureBox Picture4 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   6030
      Picture         =   "frmCompose.frx":1CB4
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   16
      Top             =   5925
      Width           =   195
   End
   Begin VB.Timer FocusTimer 
      Interval        =   1
      Left            =   255
      Top             =   3375
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00504C33&
      BorderStyle     =   0  'None
      Height          =   1260
      Left            =   105
      TabIndex        =   11
      Top             =   4860
      Width           =   4440
      Begin VB.OptionButton Option3 
         BackColor       =   &H00504C33&
         Caption         =   "Life or death"
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   75
         TabIndex        =   15
         Top             =   870
         Width           =   2790
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00504C33&
         Caption         =   "Urgent"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   75
         TabIndex        =   14
         Top             =   660
         Width           =   3900
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00504C33&
         Caption         =   "Normal"
         ForeColor       =   &H00C0C0C0&
         Height          =   210
         Left            =   75
         TabIndex        =   13
         Top             =   435
         Value           =   -1  'True
         Width           =   4290
      End
      Begin VB.Label Label3 
         BackColor       =   &H00504C33&
         Caption         =   "Priority:"
         ForeColor       =   &H00C0C0C0&
         Height          =   240
         Left            =   90
         TabIndex        =   12
         Top             =   180
         Width           =   4185
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Send"
      Height          =   495
      Left            =   4500
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5415
      Width           =   1215
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
            Picture         =   "frmCompose.frx":1FC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompose.frx":2318
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompose.frx":266A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image10 
      Height          =   3195
      Left            =   2955
      Picture         =   "frmCompose.frx":29BC
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   270
      Left            =   195
      Picture         =   "frmCompose.frx":2A02
      Stretch         =   -1  'True
      Top             =   255
      Width           =   5895
   End
   Begin VB.Image Image2 
      Height          =   5415
      Left            =   0
      Picture         =   "frmCompose.frx":2A8C
      Stretch         =   -1  'True
      Top             =   525
      Width           =   195
   End
   Begin VB.Image Image3 
      Height          =   3675
      Left            =   195
      Picture         =   "frmCompose.frx":2AF6
      Stretch         =   -1  'True
      Top             =   525
      Width           =   2925
   End
   Begin VB.Label TitleText 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Lanman Network - Compose Message"
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
      TabIndex        =   22
      Top             =   15
      Width           =   6345
   End
   Begin VB.Image Image4 
      Height          =   3675
      Left            =   3120
      Picture         =   "frmCompose.frx":2B3C
      Stretch         =   -1  'True
      Top             =   525
      Width           =   3000
   End
   Begin VB.Image Image5 
      Height          =   5460
      Left            =   6015
      Picture         =   "frmCompose.frx":2B82
      Stretch         =   -1  'True
      Top             =   525
      Width           =   195
   End
   Begin VB.Image Image6 
      Height          =   270
      Left            =   150
      Picture         =   "frmCompose.frx":2BEC
      Stretch         =   -1  'True
      Top             =   5925
      Width           =   5895
   End
   Begin VB.Image Image7 
      Height          =   255
      Left            =   360
      Top             =   0
      Width           =   5805
   End
   Begin VB.Image TopBarTile 
      Height          =   255
      Left            =   465
      Picture         =   "frmCompose.frx":2C76
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5280
   End
End
Attribute VB_Name = "frmCompose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
Dim priority As String
priority = 0
If Option1.Value = True Then
   priority = 3
End If
If Option2.Value = True Then
   priority = 2
End If
If Option3.Value = True Then
   priority = 1
End If
If priority = 0 Then
   MsgBox "Please select the priority level of this message!", vbInformation, "Messenger"
   Exit Sub
End If
If PublicOption.Value = 0 Then
   SendServerCommand "PrivateMessage" & Chr(1) & priority & Chr(1) & frmLogin.Username & Chr(1) & txtTo & Chr(1) & Subject & Chr(1) & MessageBody
Else
   SendServerCommand "PublicMessage" & Chr(1) & priority & Chr(1) & frmLogin.Username & Chr(1) & Subject & Chr(1) & MessageBody
End If
End Sub

Private Sub Image7_Click()
Dim lngReturnValue As Long
If Button = 1 Then
   Call ReleaseCapture
   lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub

Private Sub PublicOption_Click()
If PublicOption.Value = 1 Then
   txtTo.Enabled = False
   txtTo.BackColor = &H8000000F
Else
   txtTo.Enabled = True
   txtTo.BackColor = &H80000005
End If
End Sub

Private Sub Timer1_Timer()
DoEvents
If GetActiveWindow = frmCompose.hwnd Then
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
