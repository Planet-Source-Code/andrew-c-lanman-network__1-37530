VERSION 5.00
Begin VB.Form AddUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Users on LanMan Network:"
   ClientHeight    =   3480
   ClientLeft      =   1470
   ClientTop       =   1590
   ClientWidth     =   6900
   Icon            =   "AddUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox IPList 
      Height          =   2790
      Left            =   120
      TabIndex        =   5
      Top             =   3780
      Width           =   5355
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   480
      Left            =   5580
      TabIndex        =   4
      Top             =   1815
      Width           =   1215
   End
   Begin VB.CommandButton AddBtn 
      Caption         =   "&Add"
      Height          =   495
      Left            =   5580
      TabIndex        =   3
      Top             =   555
      Width           =   1215
   End
   Begin VB.ListBox UserList 
      Height          =   2790
      Left            =   120
      TabIndex        =   1
      Top             =   570
      Width           =   5355
   End
   Begin VB.CommandButton RefreshBtn 
      Caption         =   "&Refresh"
      Height          =   495
      Left            =   5580
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Click the Refresh button to search for users currently logged onto the LanMan Chat Network."
      Height          =   1065
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5160
   End
End
Attribute VB_Name = "AddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckForUsers_Timer()
    'RefreshBtn_Click
End Sub

Private Sub Command1_Click()
    AddUser.Hide
End Sub

Private Sub Form_Activate()
    'RefreshBtn_Click
End Sub

