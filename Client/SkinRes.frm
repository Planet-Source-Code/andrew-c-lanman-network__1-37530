VERSION 5.00
Begin VB.Form SkinRes 
   Caption         =   "Skin Resources"
   ClientHeight    =   6675
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9615
   Icon            =   "SkinRes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Kernel Window Resources"
      Height          =   5490
      Left            =   4905
      TabIndex        =   13
      Top             =   165
      Width           =   4155
      Begin VB.PictureBox imgkrnth 
         Height          =   660
         Left            =   75
         ScaleHeight     =   600
         ScaleWidth      =   960
         TabIndex        =   22
         Top             =   2475
         Width           =   1020
      End
      Begin VB.PictureBox imgkrnbh 
         Height          =   675
         Left            =   90
         ScaleHeight     =   615
         ScaleWidth      =   960
         TabIndex        =   21
         Top             =   3180
         Width           =   1020
      End
      Begin VB.PictureBox imgkrnwnd 
         Height          =   1320
         Left            =   1170
         ScaleHeight     =   1260
         ScaleWidth      =   2040
         TabIndex        =   20
         Top             =   2475
         Width           =   2100
      End
      Begin VB.PictureBox imgkrnlv 
         Height          =   675
         Left            =   105
         ScaleHeight     =   615
         ScaleWidth      =   930
         TabIndex        =   19
         Top             =   3900
         Width           =   990
      End
      Begin VB.PictureBox imgkrnrv 
         Height          =   690
         Left            =   1245
         ScaleHeight     =   630
         ScaleWidth      =   915
         TabIndex        =   18
         Top             =   225
         Width           =   975
      End
      Begin VB.PictureBox imgkrntr 
         Height          =   675
         Left            =   1230
         ScaleHeight     =   615
         ScaleWidth      =   930
         TabIndex        =   17
         Top             =   960
         Width           =   990
      End
      Begin VB.PictureBox imgkrntl 
         Height          =   690
         Left            =   1230
         ScaleHeight     =   630
         ScaleWidth      =   930
         TabIndex        =   16
         Top             =   1650
         Width           =   990
      End
      Begin VB.PictureBox imgkrnbr 
         Height          =   705
         Left            =   180
         ScaleHeight     =   645
         ScaleWidth      =   930
         TabIndex        =   15
         Top             =   225
         Width           =   990
      End
      Begin VB.PictureBox imgkrnbl 
         Height          =   690
         Left            =   225
         ScaleHeight     =   630
         ScaleWidth      =   870
         TabIndex        =   14
         Top             =   975
         Width           =   930
      End
      Begin VB.Shape BGColor 
         BackStyle       =   1  'Opaque
         Height          =   1305
         Left            =   1740
         Top             =   3855
         Width           =   1515
      End
   End
   Begin VB.PictureBox TopbarRight1 
      Height          =   435
      Left            =   3735
      ScaleHeight     =   375
      ScaleWidth      =   975
      TabIndex        =   12
      Top             =   60
      Width           =   1035
   End
   Begin VB.PictureBox TopBarRight0 
      Height          =   435
      Left            =   3750
      ScaleHeight     =   375
      ScaleWidth      =   975
      TabIndex        =   11
      Top             =   630
      Width           =   1035
   End
   Begin VB.PictureBox TopbarTile1 
      Height          =   615
      Left            =   2580
      ScaleHeight     =   555
      ScaleWidth      =   120
      TabIndex        =   10
      Top             =   1140
      Width           =   180
   End
   Begin VB.PictureBox TopBarTile0 
      Height          =   615
      Left            =   2850
      ScaleHeight     =   555
      ScaleWidth      =   120
      TabIndex        =   9
      Top             =   1140
      Width           =   180
   End
   Begin VB.PictureBox TopBarLeft1 
      Height          =   435
      Left            =   2550
      ScaleHeight     =   375
      ScaleWidth      =   975
      TabIndex        =   8
      Top             =   60
      Width           =   1035
   End
   Begin VB.PictureBox TopBarLeft0 
      Height          =   435
      Left            =   2580
      ScaleHeight     =   375
      ScaleWidth      =   975
      TabIndex        =   7
      Top             =   600
      Width           =   1035
   End
   Begin VB.PictureBox Button4 
      Height          =   285
      Left            =   180
      ScaleHeight     =   225
      ScaleWidth      =   4305
      TabIndex        =   6
      Top             =   3435
      Width           =   4365
   End
   Begin VB.PictureBox Button3 
      Height          =   270
      Left            =   120
      ScaleHeight     =   210
      ScaleWidth      =   4350
      TabIndex        =   5
      Top             =   2880
      Width           =   4410
   End
   Begin VB.PictureBox Button2 
      Height          =   315
      Left            =   165
      ScaleHeight     =   255
      ScaleWidth      =   4305
      TabIndex        =   4
      Top             =   2355
      Width           =   4365
   End
   Begin VB.PictureBox Button1 
      Height          =   375
      Left            =   165
      ScaleHeight     =   315
      ScaleWidth      =   4305
      TabIndex        =   3
      Top             =   1845
      Width           =   4365
   End
   Begin VB.PictureBox Button0 
      Height          =   375
      Left            =   -705
      ScaleHeight     =   315
      ScaleWidth      =   4305
      TabIndex        =   2
      Top             =   -390
      Width           =   4365
   End
   Begin VB.PictureBox MainWindow 
      Height          =   840
      Left            =   -795
      ScaleHeight     =   780
      ScaleWidth      =   3000
      TabIndex        =   1
      Top             =   -870
      Width           =   3060
   End
   Begin VB.PictureBox LoginPicture 
      Height          =   1500
      Left            =   165
      ScaleHeight     =   1440
      ScaleWidth      =   2265
      TabIndex        =   0
      Top             =   60
      Width           =   2325
   End
   Begin VB.Menu mnuGeneral 
      Caption         =   "mnuGeneral"
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogon 
         Caption         =   "Logon"
      End
      Begin VB.Menu mnuLogoff 
         Caption         =   "Logoff"
      End
      Begin VB.Menu mnuChangeUser 
         Caption         =   "Logoff and Change users"
      End
      Begin VB.Menu sep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "Status"
         Begin VB.Menu mnuSubAvailable 
            Caption         =   "Available"
         End
         Begin VB.Menu mnuSubNoSub 
            Caption         =   "Do not disturb"
         End
         Begin VB.Menu mnuSubAway 
            Caption         =   "Away"
         End
         Begin VB.Menu mnuSubOffline 
            Caption         =   "Appear offline"
         End
      End
      Begin VB.Menu sepD 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChat 
         Caption         =   "Chat"
      End
      Begin VB.Menu mnuMessage 
         Caption         =   "Message Board"
      End
      Begin VB.Menu mnuUser 
         Caption         =   "Users"
      End
   End
   Begin VB.Menu mnuSystem 
      Caption         =   "mnuSystem"
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuApp 
         Caption         =   "Application Settings"
      End
      Begin VB.Menu mnuSkin 
         Caption         =   "Skin Settings"
      End
      Begin VB.Menu mnuFServe 
         Caption         =   "FServe Settings"
      End
      Begin VB.Menu mnuConsole 
         Caption         =   "Console"
      End
   End
   Begin VB.Menu mnuUsers 
      Caption         =   "mnuUsers"
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddUser 
         Caption         =   "Add User"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove User"
      End
   End
End
Attribute VB_Name = "SkinRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuAddUser_Click()
If CheckDisabled = True Then Exit Sub
AddUser.Show
End Sub

Private Sub mnuApp_Click()
AppSettings.Show
End Sub

Private Sub mnuChangeUser_Click()
LogOff
frmLogin.Show
MessageBoard.Hide
frmCompose.Hide
MainForm.Hide
frmLogin.Password.Text = ""
frmLogin.Password.SetFocus
End Sub

Private Sub mnuConsole_Click()
frmConsole.Show
End Sub

Private Sub mnuLogoff_Click()
LogOff
End Sub

Public Sub mnuLogon_Click()
AwaitingLogin = True
MainForm.Server.Close
MainForm.Server.Connect ServerIP
'LogOn
End Sub

Private Sub mnuMessage_Click()
If CheckDisabled = True Then Exit Sub
MessageBoard.Show
End Sub


Private Sub mnuSkin_Click()
SkinCFG.Show
End Sub

Private Sub mnuSubAvailable_Click()
UserStatus = "Available"
End Sub

Private Sub mnuSubAway_Click()
UserStatus = "Away"
End Sub

Private Sub mnuSubNoSub_Click()
UserStatus = "DoNotDisturb"
End Sub

Private Sub mnuSubOffline_Click()
UserStatus = "Offline"
End Sub

Private Sub mnuUser_Click()
If CheckDisabled = True Then Exit Sub
AddUser.Show
End Sub

