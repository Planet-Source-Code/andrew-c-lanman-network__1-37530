VERSION 5.00
Begin VB.Form frmSaveFile 
   Caption         =   "Select directory"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3420
   Icon            =   "frmSaveFile.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   3420
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   480
      Left            =   1050
      TabIndex        =   2
      Top             =   3540
      Width           =   1320
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   45
      TabIndex        =   1
      Top             =   3120
      Width           =   3300
   End
   Begin VB.DirListBox Dir1 
      Height          =   3015
      Left            =   45
      TabIndex        =   0
      Top             =   75
      Width           =   3300
   End
End
Attribute VB_Name = "frmSaveFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
