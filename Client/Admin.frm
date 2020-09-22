VERSION 5.00
Begin VB.Form Admin 
   Caption         =   "LanMan Admin"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5985
   Icon            =   "Admin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton KickBtn 
      Caption         =   "Kick User"
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.ListBox UserList 
      Height          =   2400
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4515
   End
   Begin VB.Label Label1 
      Caption         =   "Users:"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
