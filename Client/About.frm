VERSION 5.00
Begin VB.Form About 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3915
   ClientLeft      =   1605
   ClientTop       =   1215
   ClientWidth     =   5655
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "About.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "About.frx":030A
   ScaleHeight     =   3915
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox CloseImage 
      AutoSize        =   -1  'True
      Height          =   795
      Left            =   2655
      Picture         =   "About.frx":4891C
      ScaleHeight     =   735
      ScaleWidth      =   1470
      TabIndex        =   0
      Top             =   3075
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   " TecTonic"
      BeginProperty Font 
         Name            =   "Matisse ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Powered by"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Image BlankImage 
      Height          =   735
      Left            =   180
      Top             =   3060
      Width           =   1470
   End
   Begin VB.Image CloseBTN 
      Height          =   735
      Left            =   4155
      Top             =   3105
      Width           =   1470
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CloseBTN_Click()
Me.Hide
End Sub

Private Sub CloseBTN_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
CloseBTN.Picture = CloseImage.Picture
End Sub

Private Sub CloseBTN_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
CloseBTN.Picture = BlankImage.Picture
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    CloseBTN.Picture = CloseImage.Picture
    DoEvents
    Me.Hide
End If
End Sub

