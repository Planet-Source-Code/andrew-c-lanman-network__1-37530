VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmFServe 
   Caption         =   "Lanman File Server"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6840
   Icon            =   "frmFServe.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Outgoing 
      Index           =   0
      Left            =   5430
      Top             =   1230
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsClient 
      Index           =   0
      Left            =   5910
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Server 
      Left            =   5430
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   6169
   End
   Begin VB.CommandButton LocalCMD 
      Caption         =   "Local Files"
      Height          =   420
      Left            =   4005
      TabIndex        =   10
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton UploadsBTN 
      Caption         =   "Uploads"
      Height          =   420
      Left            =   2670
      TabIndex        =   3
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton DownloadsBTN 
      Caption         =   "Downloads"
      Height          =   420
      Left            =   1335
      TabIndex        =   2
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton SearchBTN 
      Caption         =   "Search"
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6075
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame SearchBar 
      Height          =   495
      Left            =   -15
      TabIndex        =   4
      Top             =   315
      Width           =   6870
      Begin VB.CommandButton cmdGo 
         Caption         =   "&Go"
         Default         =   -1  'True
         Height          =   315
         Left            =   6075
         TabIndex        =   6
         Top             =   135
         Width           =   720
      End
      Begin VB.TextBox SearchText 
         Height          =   315
         Left            =   30
         TabIndex        =   5
         Text            =   "Enter search filename/keyword"
         Top             =   135
         Width           =   6015
      End
   End
   Begin VB.Frame DownloadsFrame 
      Height          =   5400
      Left            =   -15
      TabIndex        =   14
      Top             =   690
      Width           =   6870
      Begin VB.CommandButton Command4 
         Caption         =   "Clear"
         Height          =   285
         Left            =   90
         TabIndex        =   22
         Top             =   5025
         Width           =   1185
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2325
         Left            =   75
         TabIndex        =   21
         Top             =   2685
         Width           =   6705
         _ExtentX        =   11827
         _ExtentY        =   4101
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Filename"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Filesize"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Source"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Download Time"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Resume"
         Height          =   300
         Left            =   2505
         TabIndex        =   19
         Top             =   2055
         Width           =   1200
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Pause"
         Height          =   300
         Left            =   1290
         TabIndex        =   18
         Top             =   2055
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cancel"
         Height          =   300
         Left            =   90
         TabIndex        =   17
         Top             =   2055
         Width           =   1200
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1635
         Left            =   75
         TabIndex        =   16
         Top             =   405
         Width           =   6720
         _ExtentX        =   11853
         _ExtentY        =   2884
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Filename"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Filesize"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Source"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Percent"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Speed"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label4 
         Caption         =   "Complete Downloads:"
         Height          =   240
         Left            =   75
         TabIndex        =   20
         Top             =   2475
         Width           =   5655
      End
      Begin VB.Label Label3 
         Caption         =   "Current Downloads:"
         Height          =   225
         Left            =   90
         TabIndex        =   15
         Top             =   195
         Width           =   5325
      End
   End
   Begin VB.Frame SearchFrame 
      Height          =   5400
      Left            =   -15
      TabIndex        =   7
      Top             =   690
      Width           =   6870
      Begin VB.CheckBox LockDir 
         Caption         =   "Lock download directory"
         Height          =   210
         Left            =   1365
         TabIndex        =   24
         Top             =   5040
         Width           =   4695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Download"
         Height          =   360
         Left            =   105
         TabIndex        =   23
         Top             =   4935
         Width           =   1155
      End
      Begin MSComctlLib.ListView SearchList 
         Height          =   4485
         Left            =   75
         TabIndex        =   9
         Top             =   435
         Width           =   6690
         _ExtentX        =   11800
         _ExtentY        =   7911
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Filename"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Filesize (Bytes)"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Host"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Search Results:"
         Height          =   225
         Left            =   90
         TabIndex        =   8
         Top             =   195
         Width           =   5325
      End
   End
   Begin VB.Frame LocalFrame 
      Height          =   5400
      Left            =   -15
      TabIndex        =   11
      Top             =   690
      Width           =   6870
      Begin VB.FileListBox LocalFiles 
         Height          =   4770
         Left            =   75
         TabIndex        =   13
         Top             =   435
         Width           =   6510
      End
      Begin VB.Label Label2 
         Caption         =   "Local shared files:"
         Height          =   225
         Left            =   90
         TabIndex        =   12
         Top             =   195
         Width           =   5325
      End
   End
End
Attribute VB_Name = "frmFServe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Searching As Boolean

Private Sub cmdGo_Click()
If cmdGo.Caption = "&Stop" Then
   If Searching = True Then
      SendServerCommand "KillSearch"
      Searching = False
      cmdGo.Caption = "&Go"
      Exit Sub
   End If
Else
   Sleep 500
   cmdGo.Caption = "&Stop"
   SearchList.ListItems.Clear
   SendServerCommand "InitSearch" & Chr(1) & SearchText.Text
   Me.Caption = "Lanman File Server - Searching..."
   StatusBar1.SimpleText = "Searching..."
   Searching = True
End If
End Sub

Private Sub DownloadsBTN_Click()
DownloadsFrame.ZOrder (0)
SearchBar.ZOrder (0)
SearchBTN.ZOrder (0)
DownloadsBTN.ZOrder (0)
UploadsBTN.ZOrder (0)
LocalCMD.ZOrder (0)
End Sub

Private Sub Form_Load()
On Error Resume Next
MkDir "C:\Shared Folder" 'Create it just in case.
LocalFiles.Path = "C:\Shared Folder"
End Sub

Private Sub LocalCMD_Click()
LocalFrame.ZOrder (0)
SearchBar.ZOrder (0)
SearchBTN.ZOrder (0)
DownloadsBTN.ZOrder (0)
UploadsBTN.ZOrder (0)
LocalCMD.ZOrder (0)
End Sub

Private Sub SearchBTN_Click()
SearchFrame.ZOrder (0)
SearchBar.ZOrder (0)
SearchBTN.ZOrder (0)
DownloadsBTN.ZOrder (0)
UploadsBTN.ZOrder (0)
LocalCMD.ZOrder (0)
End Sub
