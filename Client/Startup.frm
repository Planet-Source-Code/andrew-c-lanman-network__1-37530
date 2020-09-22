VERSION 5.00
Begin VB.Form Startup 
   BorderStyle     =   0  'None
   ClientHeight    =   2565
   ClientLeft      =   1935
   ClientTop       =   2115
   ClientWidth     =   4560
   ControlBox      =   0   'False
   Icon            =   "Startup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   Picture         =   "Startup.frx":030A
   ScaleHeight     =   2565
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Settings 
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   1695
      Picture         =   "Startup.frx":26ABC
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox Forms 
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   1215
      Picture         =   "Startup.frx":2742E
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox Socks 
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   735
      Picture         =   "Startup.frx":27DA0
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox Cards 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   240
      Picture         =   "Startup.frx":28712
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label Done3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "- Done"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   3915
      TabIndex        =   11
      Top             =   1740
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Done4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "- Done"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   3915
      TabIndex        =   10
      Top             =   2025
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Done2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "- Done"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   3915
      TabIndex        =   9
      Top             =   1470
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Done1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "- Done"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   3915
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Status4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "{Loading Text}"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   240
      TabIndex        =   7
      Top             =   2010
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label Status3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "{Loading Text}"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   240
      TabIndex        =   6
      Top             =   1740
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label Status2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "{Loading Text}"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   240
      TabIndex        =   5
      Top             =   1470
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label Status1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "{Loading Text}"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   240
      TabIndex        =   4
      Top             =   1215
      Visible         =   0   'False
      Width           =   1245
   End
End
Attribute VB_Name = "Startup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// Startup form...
'// Initialize
Public OKforLogon As Boolean

Public Sub Initialize()
Cards.Visible = False
Socks.Visible = False
Forms.Visible = False
Settings.Visible = False
Status1.Visible = False
Status2.Visible = False
Status3.Visible = False
Status4.Visible = False
Done1.Visible = False
Done2.Visible = False
Done3.Visible = False
Done4.Visible = False

Startup.Show
Randomize
Do Until DoEvents
Loop
Status1.Visible = True
Cards.Visible = True
Status1.Caption = "Loading previous settings..."
InitSettings
DoEvents
Done1.Visible = True

Status2.Visible = True
Socks.Visible = True
Status2.Caption = "Loading skin data..."
LoadSkinData GetKey("Lanman", "Settings", "Skin", "Default")
DoEvents
Sleep Int(Rnd * 200) + 200
Done2.Visible = True

Status3.Visible = True
Forms.Visible = True
Status3.Caption = "Loading forms into memory..."
InitForms
DoEvents
Done3.Visible = True

Status4.Visible = True
Settings.Visible = True
If frmLogin.Check1.Value = 1 Then
   Status4.Caption = "Authenticating with server..."
Else
   Status4.Caption = "Awaiting user login..."
End If
Sleep 500
If frmLogin.Check1.Value = 1 Then
    frmLogin.Username.Text = GetKey("Lanman", "Settings", "Username", "Guest" & Int(Rnd * 65535))
    frmLogin.Password.Text = GetKey("Lanman", "Settings", "Password", "")
    frmLogin.AwaitingLogin = True
    MainForm.Server.Close
    MainForm.Server.Connect ServerIP
Else
    Randomize
    frmLogin.Username.Text = GetKey("Lanman", "Settings", "Username", "Guest" & Int(Rnd * 65535))
    frmLogin.Show
    Startup.Hide
    MainForm.Hide
    UserStatus = "Offline"
End If
Me.Hide
End Sub

