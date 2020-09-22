VERSION 5.00
Begin VB.Form Popup 
   BackColor       =   &H00FF8080&
   ClientHeight    =   2100
   ClientLeft      =   5655
   ClientTop       =   5820
   ClientWidth     =   2610
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "Popup.frx":0000
   ScaleHeight     =   2100
   ScaleWidth      =   2610
   ShowInTaskbar   =   0   'False
   Begin VB.Timer WindowActive 
      Interval        =   1
      Left            =   120
      Top             =   1080
   End
   Begin VB.Timer MouseOver 
      Interval        =   1000
      Left            =   120
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   2100
   End
   Begin VB.Label Message 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1515
      Left            =   90
      TabIndex        =   1
      Top             =   450
      Width           =   2445
      WordWrap        =   -1  'True
   End
   Begin VB.Label Title 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2370
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Popup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public numSec As Integer
Public OrigHeight As Long
Public StartMenuAdj As Long
Public sMouseOver As Boolean
Public ToggleChange As Boolean
Public ResourceLoaded As Boolean
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
    End Type
Private What As RECT

Private Sub Form_Load()
If ResourceLoaded = False Then
   Popup.Picture = SkinRes.PopupImage.Picture
   ResourceLoaded = True
End If
numsecs = 0
Popup.Width = 2625
StartMenuAdj = GetStartMenuHeight
OrigHeight = Popup.Height
Popup.Top = GetStartMenuHeight + 50
Popup.Left = Screen.Width - Popup.Width - 10
Popup.Height = 0
OpenForm Me, 5
Timer1.Interval = 1000
End Sub

Public Function closeForm(frm As Form, speed As Integer)
Do Until Popup.Height = 120
    Dim X
    X = GetStartMenuHeight / 1440
    DoEvents
    Popup.Height = Popup.Height - speed * 5
    Popup.Top = Popup.Top + speed * 5
Loop
numsecs = 0
Unload frm
End Function
Public Function OpenForm(closeForm As Form, speed As Integer)
Timer1.Enabled = True
closeForm.Show
closeForm.ScaleMode = 1
closeForm.WindowState = 0
Do Until closeForm.Height >= OrigHeight
Dim X
X = GetStartMenuHeight / 1440
DoEvents
closeForm.Height = closeForm.Height + speed * 9
closeForm.Top = closeForm.Top - speed * 9
Loop
numsecs = 0
End Function

Private Sub MouseOver_Timer()
Dim CursorPOS As POINTAPI
GetCursorPos CursorPOS
If CursorPOS.X * Screen.TwipsPerPixelX > Popup.Left And CursorPOS.Y * Screen.TwipsPerPixelY > Popup.Top Then
        Message.ForeColor = RGB(0, 0, 255)
        Message.Font.Underline = True
        sMouseOver = True
Else
    Message.ForeColor = RGB(0, 0, 0)
    Message.Font.Bold = False
    Message.Font.Underline = False
    sMouseOver = False
End If
End Sub

Private Sub Timer1_Timer()
numSec = numSec + 1
If numSec >= 4 Then
    closeForm Me, 15
End If
End Sub

Private Sub WindowActive_Timer()
Dim CursorPOS As POINTAPI
GetCursorPos CursorPOS
If CursorPOS.X * Screen.TwipsPerPixelX > Popup.Left Then
    If CursorPOS.Y * Screen.TwipsPerPixelY > Popup.Top Then
        sMouseOver = True
    Else
        sMouseOver = False
    End If
Else
    sMouseOver = False
End If
If sMouseOver = True Then
   numSec = 0
End If
End Sub
