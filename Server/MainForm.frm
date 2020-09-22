VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form MainForm 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lanman Chat Server"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   6705
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock wsClient 
      Index           =   0
      Left            =   825
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsServer 
      Left            =   345
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   6168
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4755
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   8387
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      Tab             =   2
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "MainForm.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdStartStop"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Users"
      TabPicture(1)   =   "MainForm.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "UserList"
      Tab(1).Control(1)=   "Command1"
      Tab(1).Control(2)=   "Command2"
      Tab(1).Control(3)=   "Command3"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Rooms"
      TabPicture(2)   =   "MainForm.frx":0044
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Command4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Command5"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Command6"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Roomlist"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Frame3"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Message Board"
      TabPicture(3)   =   "MainForm.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "TempDir"
      Tab(3).Control(1)=   "TempFile"
      Tab(3).Control(2)=   "Frame1"
      Tab(3).Control(3)=   "lblPrivateMSGCount"
      Tab(3).Control(4)=   "lblPublicMSGCount"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "File Server"
      TabPicture(4)   =   "MainForm.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "FileList"
      Tab(4).Control(1)=   "Frame2"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "About"
      TabPicture(5)   =   "MainForm.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      Begin VB.Frame Frame3 
         Caption         =   "Show Only:"
         Height          =   1695
         Left            =   5340
         TabIndex        =   27
         Top             =   2070
         Width           =   1290
         Begin VB.OptionButton Option3 
            Caption         =   "Both"
            Height          =   225
            Left            =   90
            TabIndex        =   30
            Top             =   825
            Value           =   -1  'True
            Width           =   1080
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Private"
            Height          =   210
            Left            =   90
            TabIndex        =   29
            Top             =   540
            Width           =   1110
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Public"
            Height          =   195
            Left            =   90
            TabIndex        =   28
            Top             =   255
            Width           =   1095
         End
      End
      Begin MSComctlLib.ListView Roomlist 
         Height          =   4185
         Left            =   75
         TabIndex        =   26
         Top             =   465
         Width           =   5205
         _ExtentX        =   9181
         _ExtentY        =   7382
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Room name"
            Object.Width           =   5644
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Users"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Type"
            Object.Width           =   1587
         EndProperty
      End
      Begin MSComctlLib.ListView FileList 
         Height          =   1650
         Left            =   -74910
         TabIndex        =   25
         Top             =   2955
         Width           =   6480
         _ExtentX        =   11430
         _ExtentY        =   2910
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "File Size"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Host"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame Frame2 
         Caption         =   "Statistics:"
         Height          =   2580
         Left            =   -74910
         TabIndex        =   23
         Top             =   375
         Width           =   6480
         Begin VB.Label lblTotalFiles 
            Caption         =   "Total Network Capacity:"
            Height          =   240
            Left            =   150
            TabIndex        =   24
            Top             =   255
            Width           =   6000
         End
      End
      Begin VB.DirListBox TempDir 
         Height          =   315
         Left            =   -72360
         TabIndex        =   22
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.FileListBox TempFile 
         Height          =   480
         Left            =   -73440
         TabIndex        =   21
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   12
         Top             =   960
         Width           =   6375
         Begin VB.CommandButton Command9 
            Caption         =   "Print"
            Height          =   375
            Left            =   5040
            TabIndex        =   20
            Top             =   3000
            Width           =   975
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Delete"
            Height          =   375
            Left            =   3960
            TabIndex        =   19
            Top             =   3000
            Width           =   975
         End
         Begin VB.CommandButton Command7 
            Caption         =   "View"
            Height          =   375
            Left            =   2760
            TabIndex        =   18
            Top             =   3000
            Width           =   1095
         End
         Begin VB.ListBox pMessageList 
            Height          =   2400
            Left            =   2520
            TabIndex        =   17
            Top             =   480
            Width           =   3735
         End
         Begin VB.DirListBox pInboxDir 
            Height          =   1215
            Left            =   240
            TabIndex        =   15
            Top             =   2160
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.ListBox pInboxList 
            Height          =   2985
            Left            =   120
            TabIndex        =   14
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label Label2 
            Caption         =   "Contents of inbox:"
            Height          =   255
            Left            =   2520
            TabIndex        =   16
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Private inboxes:"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   3255
         End
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Create Private"
         Height          =   465
         Left            =   5325
         TabIndex        =   9
         Top             =   1530
         Width           =   1305
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Create Public"
         Height          =   465
         Left            =   5325
         TabIndex        =   8
         Top             =   1005
         Width           =   1305
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Close Room"
         Height          =   465
         Left            =   5325
         TabIndex        =   7
         Top             =   480
         Width           =   1305
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Whois"
         Height          =   465
         Left            =   -69675
         TabIndex        =   6
         Top             =   1530
         Width           =   1305
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ban"
         Height          =   465
         Left            =   -69675
         TabIndex        =   5
         Top             =   1005
         Width           =   1305
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Kick"
         Height          =   465
         Left            =   -69675
         TabIndex        =   4
         Top             =   480
         Width           =   1305
      End
      Begin MSComctlLib.ListView UserList 
         Height          =   4185
         Left            =   -74925
         TabIndex        =   3
         Top             =   465
         Width           =   5205
         _ExtentX        =   9181
         _ExtentY        =   7382
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Username"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Socket"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Search Status"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmdStartStop 
         Caption         =   "&Stop Server"
         Height          =   540
         Left            =   -74910
         TabIndex        =   2
         Top             =   4110
         Width           =   1440
      End
      Begin VB.Label lblPrivateMSGCount 
         Caption         =   "Private Messages:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   11
         Top             =   720
         Width           =   4575
      End
      Begin VB.Label lblPublicMSGCount 
         Caption         =   "Public Messages:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   10
         Top             =   480
         Width           =   3975
      End
   End
   Begin VB.FileListBox AccountList 
      Height          =   2625
      Left            =   7800
      TabIndex        =   0
      Top             =   6240
      Visible         =   0   'False
      Width           =   4095
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ActivityClose_Click()
ActivityFrame.Visible = False
End Sub

Private Sub cmdStartStop_Click()
If cmdStartStop.Caption = "&Start Server" Then
   cmdStartStop.Caption = "&Stop Server"
   StartServer
Else
   cmdStartStop.Caption = "&Start Server"
   StopServer
End If
End Sub

Private Sub Command5_Click()
ChatRooms.CreateRoom InputBox("Room name:", "Create public room"), "Public", -1
End Sub

Private Sub Command6_Click()
ChatRooms.CreateRoom InputBox("Room name:", "Create private room"), "Private", -1
End Sub

Private Sub Form_Load()
CheckPaths
UpdateCounts
UpdateInboxs
StartServer
End Sub

Private Sub Label2_Click()
If ActivityFrame.Visible = True Then
    ActivityFrame.Visible = False
Else
    ActivityFrame.Visible = True
    UsersFrame.Visible = False
End If
End Sub

Private Sub Label3_Click()
If UsersFrame.Visible = True Then
    UsersFrame.Visible = False
Else
    UsersFrame.Visible = True
    ActivityFrame.Visible = False
End If
End Sub

Private Sub UsersClose_Click()
UsersFrame.Visible = False
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
   RefreshRoomList "Public"
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
   RefreshRoomList "Private"
End If
End Sub

Private Sub Option3_Click()
If Option3.Value = True Then ShowAllRooms
End Sub

Private Sub pInboxList_Click()
If pInboxList.Text <> "" Then
   TempFile.Path = App.Path & "\Messages\" & pInboxList.Text
   For x = 0 To TempFile.ListCount - 1
   If InStrRev(TempFile.List(x), ".", Len(TempFile.List(x)), vbTextCompare) = 0 Then
      pMessageList.AddItem TempFile.List(x)
   Else
      pMessageList.AddItem Left(TempFile.List(x), InStrRev(TempFile.List(x), ".", Len(TempFile.List(x)), vbTextCompare) - 1)
   End If
   Next x
End If
If pMessageList.ListCount = 0 Then
   pMessageList.AddItem "Inbox Empty"
End If
End Sub

Private Sub Roomlist_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
   Roomlist.StartLabelEdit
End If
End Sub

Private Sub wsClient_Close(Index As Integer)
On Error Resume Next
'// Search through and remove user from all chatrooms
For x = 0 To UBound(RoomNames)
If IsUserInRoom(Right(RoomNames(x), Len(RoomNames(x)) - 2), GetUsernameFromWS(Index)) = True Then
   RemoveFromRoom (Right(RoomNames(x), Len(RoomNames(x)) - 2)), GetUsernameFromWS(Index)
End If
Next x
wsClient(Index).Close
UserList.ListItems.Remove LocateUserFromWS(Index)
End Sub

Private Sub wsClient_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strData As String
wsClient(Index).GetData strData
ProcessCommand strData, Index
End Sub

Private Sub wsServer_ConnectionRequest(ByVal requestID As Long)
wsClient(FindFreeSock).Accept requestID
End Sub
