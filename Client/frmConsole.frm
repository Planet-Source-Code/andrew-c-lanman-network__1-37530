VERSION 5.00
Begin VB.Form frmConsole 
   BackColor       =   &H00404000&
   BorderStyle     =   0  'None
   ClientHeight    =   4935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   165
      Left            =   90
      TabIndex        =   8
      Text            =   "Lanman:>"
      Top             =   4680
      Width           =   4290
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   4305
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "frmConsole.frx":0000
      Top             =   375
      Width           =   4290
   End
   Begin VB.PictureBox TopBarLeft 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      Picture         =   "frmConsole.frx":002F
      ScaleHeight     =   255
      ScaleWidth      =   465
      TabIndex        =   5
      Top             =   0
      Width           =   465
   End
   Begin VB.PictureBox TopBarRight 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3780
      Picture         =   "frmConsole.frx":06D1
      ScaleHeight     =   255
      ScaleWidth      =   705
      TabIndex        =   4
      Top             =   0
      Width           =   705
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   0
      Picture         =   "frmConsole.frx":10A3
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
      Picture         =   "frmConsole.frx":13B5
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   4680
      Width           =   195
   End
   Begin VB.PictureBox Picture4 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   4305
      Picture         =   "frmConsole.frx":16C7
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   4680
      Width           =   195
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   4290
      Picture         =   "frmConsole.frx":19D9
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   255
      Width           =   195
   End
   Begin VB.Image Image4 
      Height          =   4275
      Left            =   960
      Picture         =   "frmConsole.frx":1CEB
      Stretch         =   -1  'True
      Top             =   480
      Width           =   3360
   End
   Begin VB.Label TitleText 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Command Console"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   1545
      TabIndex        =   6
      Top             =   15
      Width           =   1365
   End
   Begin VB.Image TopBarTile 
      Height          =   255
      Left            =   465
      Picture         =   "frmConsole.frx":1D31
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3810
   End
   Begin VB.Image Image1 
      Height          =   270
      Left            =   195
      Picture         =   "frmConsole.frx":1DB7
      Stretch         =   -1  'True
      Top             =   255
      Width           =   4110
   End
   Begin VB.Image Image2 
      Height          =   4155
      Left            =   0
      Picture         =   "frmConsole.frx":1E41
      Stretch         =   -1  'True
      Top             =   525
      Width           =   195
   End
   Begin VB.Image Image3 
      Height          =   4275
      Left            =   165
      Picture         =   "frmConsole.frx":1EAB
      Stretch         =   -1  'True
      Top             =   465
      Width           =   3360
   End
   Begin VB.Image Image5 
      Height          =   4155
      Left            =   4290
      Picture         =   "frmConsole.frx":1EF1
      Stretch         =   -1  'True
      Top             =   525
      Width           =   195
   End
   Begin VB.Image Image6 
      Height          =   270
      Left            =   195
      Picture         =   "frmConsole.frx":1F5B
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   4110
   End
End
Attribute VB_Name = "frmConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub
