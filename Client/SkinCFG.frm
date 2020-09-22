VERSION 5.00
Begin VB.Form SkinCFG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lanman Skin Settings"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   Icon            =   "SkinCFG.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox AutoPreview 
      Caption         =   "Auto preview"
      Height          =   225
      Left            =   5520
      TabIndex        =   10
      Top             =   2025
      Width           =   1245
   End
   Begin VB.DirListBox SkinDir 
      Height          =   1890
      Left            =   600
      TabIndex        =   9
      Top             =   1680
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      Caption         =   "Skin Details:"
      Height          =   1575
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   5295
      Begin VB.Label lblDescription 
         Caption         =   "Description:"
         Height          =   735
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   5055
      End
      Begin VB.Label lblAuthor 
         Caption         =   "Author:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label lblTitle 
         Caption         =   "Title:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Available Skins:"
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5295
      Begin VB.ListBox SkinList 
         Height          =   1620
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Preview"
      Height          =   495
      Left            =   5520
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   495
      Left            =   5520
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Apply Skin"
      Height          =   495
      Left            =   5520
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "SkinCFG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SkinChanged As Boolean
Private Sub Command1_Click()
On Error GoTo NoSkin
Dim SkinName As String
Dim sFilename As String
For X = 0 To SkinList.ListCount
SkinName = GetINIVal("SkinInfo", "Title", SkinDir.List(X) & "\skin.dat")
If SkinName = SkinList.List(SkinList.ListIndex) Then
   sFilename = Right(SkinDir.List(X), Len(SkinDir.List(X)) - InStrRev(SkinDir.List(X), "\", Len(SkinDir.List(X)), vbTextCompare))
   GoTo FoundValue
End If
Next X
FoundValue:
LoadSkinData sFilename
SaveKey "Lanman", "Settings", "Skin", sFilename
Me.Hide
NoSkin:
End Sub

Private Sub Command2_Click()
If SkinChanged = True Then
    LoadSkinData GetKey("Lanman", "Settings", "Skin", "Default")
End If
Me.Hide
End Sub

Private Sub Command3_Click()
Dim SkinName As String
Dim sFilename As String
For X = 0 To SkinList.ListCount
SkinName = GetINIVal("SkinInfo", "Title", SkinDir.List(X) & "\skin.dat")
If SkinName = SkinList.List(SkinList.ListIndex) Then
   sFilename = Right(SkinDir.List(X), Len(SkinDir.List(X)) - InStrRev(SkinDir.List(X), "\", Len(SkinDir.List(X)), vbTextCompare))
   GoTo FoundValue
End If
Next X
FoundValue:
LoadSkinData sFilename
SkinChanged = True
End Sub

Private Sub Form_Load()
SkinDir.Path = App.Path & "\Skinz"
SkinList.Clear
For X = 0 To SkinDir.ListCount
Dim SkinName
SkinName = GetINIVal("SkinInfo", "Title", SkinDir.List(X) & "\skin.dat")
SkinList.AddItem SkinName
Next X
End Sub

Private Sub LoadSkinDetails(SkinName As String)
Dim sSkinName
Dim Author
Dim Description
sSkinName = GetINIVal("SkinInfo", "Title", App.Path & "\Skinz\" & SkinName & "\skin.dat")
Author = GetINIVal("SkinInfo", "Author", App.Path & "\Skinz\" & SkinName & "\skin.dat")
Description = GetINIVal("SkinInfo", "Description", App.Path & "\Skinz\" & SkinName & "\skin.dat")
SkinCFG.lblTitle.Caption = sSkinName
SkinCFG.lblAuthor.Caption = Author
SkinCFG.lblDescription.Caption = Description
End Sub

Private Sub SkinList_Click()
Dim SkinName
Dim Author
Dim Description
Dim Filename
For X = 0 To SkinList.ListCount
SkinName = GetINIVal("SkinInfo", "Title", SkinDir.List(X) & "\skin.dat")
Author = GetINIVal("SkinInfo", "Author", SkinDir.List(X) & "\skin.dat")
Description = GetINIVal("SkinInfo", "Description", SkinDir.List(X) & "\skin.dat")
If SkinName = SkinList.List(SkinList.ListIndex) Then
   Filename = SkinDir.List(X)
   GoTo FoundValue
End If
Next X
FoundValue:
lblTitle.Caption = "Title: " & SkinName
lblAuthor.Caption = "Author: " & Author
lblDescription.Caption = "Description: " & Description
If AutoPreview.Value = 1 Then
   Call Command3_Click
End If
End Sub
