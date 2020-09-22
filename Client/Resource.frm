VERSION 5.00
Begin VB.Form Resource 
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   1650
   ClientTop       =   1830
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.Label onlinestatus 
      Caption         =   "false"
      Height          =   360
      Left            =   105
      TabIndex        =   2
      Top             =   870
      Width           =   1545
   End
   Begin VB.Label DestName 
      Caption         =   "Label1"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   495
      Width           =   4815
   End
   Begin VB.Label SettingsFirst 
      Caption         =   "no"
      Height          =   255
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   5145
   End
   Begin VB.Menu usermenuParent 
      Caption         =   "UserMenu"
      Begin VB.Menu mnuPopup 
         Caption         =   "Send Popup Message"
      End
   End
End
Attribute VB_Name = "Resource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuAbout_Click()
About.Show
End Sub

Private Sub mnuExit_Click()
SaveKey "Lanman", "Settings", "XPos", MainForm.Left
SaveKey "Lanman", "Settings", "YPos", MainForm.Top
End
End Sub

Private Sub mnuLogoff_Click()
LogOff
End Sub

Private Sub mnuLogon_Click()
LogOn
End Sub

Private Sub SettingsMenu_Click()
AppSettings.Show
End Sub

Private Sub ChatWndOpen_Click()

End Sub

