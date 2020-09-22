VERSION 5.00
Begin VB.Form ChatRequest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chat request"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   Icon            =   "ChatRequest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Deny All"
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Deny"
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Accept"
      Default         =   -1  'True
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label WinsockIndex 
      Caption         =   "0"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "User is requesting to chat. Do you accept?"
      Height          =   570
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   5070
   End
End
Attribute VB_Name = "ChatRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
