VERSION 5.00
Begin VB.Form UserDetails 
   Caption         =   "User Details - [username]"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5745
   Icon            =   "UserDetails.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   5745
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Top             =   600
      Width           =   3375
   End
   Begin VB.CommandButton AddBtn 
      Caption         =   "&Add"
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Nametext 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Email:"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Full Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "UserDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
