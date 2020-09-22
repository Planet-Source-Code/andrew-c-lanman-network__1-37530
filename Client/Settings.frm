VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form AppSettings 
   Caption         =   "Program Settings"
   ClientHeight    =   5640
   ClientLeft      =   1920
   ClientTop       =   3615
   ClientWidth     =   7620
   Icon            =   "Settings.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Import"
      Height          =   495
      Left            =   1545
      TabIndex        =   9
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Export"
      Height          =   495
      Left            =   3105
      TabIndex        =   8
      Top             =   5040
      Width           =   1425
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   4665
      TabIndex        =   0
      Top             =   5040
      Width           =   1425
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   8493
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "Settings.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Picture1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Network"
      TabPicture(1)   =   "Settings.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Picture2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "File Sharing"
      TabPicture(2)   =   "Settings.frx":0342
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Picture3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Security"
      TabPicture(3)   =   "Settings.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Picture4"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.PictureBox Picture4 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   4350
         Left            =   -74880
         Picture         =   "Settings.frx":037A
         ScaleHeight     =   4350
         ScaleWidth      =   7215
         TabIndex        =   13
         Top             =   360
         Width           =   7215
      End
      Begin VB.PictureBox Picture3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   4350
         Left            =   120
         Picture         =   "Settings.frx":66784
         ScaleHeight     =   4350
         ScaleWidth      =   7215
         TabIndex        =   12
         Top             =   360
         Width           =   7215
         Begin VB.Frame Frame3 
            Caption         =   "Shared Directories:"
            Height          =   2565
            Left            =   0
            TabIndex        =   17
            Top             =   1440
            Width           =   7095
            Begin VB.ListBox SharedList 
               Height          =   2205
               Left            =   120
               TabIndex        =   20
               Top             =   225
               Width           =   6045
            End
            Begin VB.CommandButton Command5 
               Caption         =   "Remove"
               Height          =   300
               Left            =   6255
               TabIndex        =   19
               Top             =   585
               Width           =   735
            End
            Begin VB.CommandButton Command4 
               Caption         =   "Add"
               Height          =   300
               Left            =   6255
               TabIndex        =   18
               Top             =   225
               Width           =   735
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "File sharing preferences:"
            Height          =   1095
            Left            =   0
            TabIndex        =   14
            Top             =   120
            Width           =   7095
            Begin VB.CheckBox DownloadWarning 
               Caption         =   "Warn when someone attempts download"
               Height          =   255
               Left            =   120
               TabIndex        =   16
               Top             =   600
               Width           =   6855
            End
            Begin VB.CheckBox FileServerEnable 
               Caption         =   "Enable file sharing"
               Height          =   255
               Left            =   120
               TabIndex        =   15
               Top             =   240
               Value           =   1  'Checked
               Width           =   6855
            End
         End
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   4350
         Left            =   -74880
         Picture         =   "Settings.frx":CCB8E
         ScaleHeight     =   4350
         ScaleWidth      =   7215
         TabIndex        =   11
         Top             =   360
         Width           =   7215
      End
      Begin VB.Frame Frame2 
         Caption         =   "Application Preferences:"
         Height          =   1815
         Left            =   -74880
         TabIndex        =   3
         Top             =   480
         Width           =   7095
         Begin VB.CheckBox ShowStartup 
            Caption         =   "Show startup screen while initializing"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   1320
            Width           =   5055
         End
         Begin VB.CheckBox Startmin 
            Caption         =   "Start minimised in system-tray"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   960
            Width           =   5295
         End
         Begin VB.CheckBox AutoLogon 
            Caption         =   "Remember username and password and log-in automatically"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   5655
         End
         Begin VB.CheckBox AutoStart 
            Caption         =   "Automatically start with windows"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   5535
         End
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   4350
         Left            =   -74880
         Picture         =   "Settings.frx":132F98
         ScaleHeight     =   4350
         ScaleWidth      =   7215
         TabIndex        =   10
         Top             =   360
         Width           =   7215
      End
      Begin VB.Label Label1 
         Caption         =   "*This distribution is configured to connect to a static server ip. You cannot change the remote server."
         Height          =   735
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   7095
      End
   End
End
Attribute VB_Name = "AppSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AboutBtn_Click()
About.Show
End Sub

'// Settings form...
Private Sub Command1_Click()
'// Save all settings into registry.
On Error Resume Next
If AutoStart.Value = 0 Then
   FileSystem.Kill "C:\windows\startm~1\programs\startup\lanman~1.lnk"
Else
   CreateShellLink "C:\program files\lanman\lanman.exe", "startup", "", "Lanman Startup", True, "$(Programs)"
End If
SaveKey "Lanman", "Settings", "Autostart", AutoStart.Value
SaveKey "Lanman", "Settings", "AutoLogon", AutoLogon.Value
SaveKey "Lanman", "Settings", "StartMin", Startmin.Value
SaveKey "Lanman", "Settings", "ShowStartup", ShowStartup.Value
If AutoLogon.Value = 1 Then
   SaveKey "Lanman", "Settings", "Password", frmLogin.Password.Text
   SaveKey "Lanman", "Settings", "Username", frmLogin.Username.Text
   frmLogin.Check1.Value = 1
End If
FirstTimeMessage.Visible = False
AppSettings.Hide
End Sub

Private Sub DefaultBtn_Click()
'// Load/Set default settings
AutoStart.Value = 1
AutoLogon.Value = 0
Startmin.Value = 0
SaveKey "Lanman", "Settings", "Autostart", AutoStart.Value
SaveKey "Lanman", "Settings", "AutoLogon", AutoLogon.Value
SaveKey "Lanman", "Settings", "StartMin", Startmin.Value
End Sub

