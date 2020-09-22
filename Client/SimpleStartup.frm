VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form SimpleStartup 
   Caption         =   "LanMan - Loading..."
   ClientHeight    =   1335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5910
   Icon            =   "SimpleStartup.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1335
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   2175
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "LanMan is starting up...Please wait..."
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "SimpleStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
TerminateForms
End
End Sub

Private Sub Form_Load()
LoadEverything
End Sub

Private Sub Timer1_Timer()
On Error GoTo endofproc
ProgressBar1.Value = ProgressBar1.Value + 10
Exit Sub
endofproc:
SimpleStartup.Hide
MainForm.Show

End Sub
