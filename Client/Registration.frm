VERSION 5.00
Begin VB.Form Registration 
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   6045
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2160
   ScaleWidth      =   6045
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   3405
      Left            =   105
      TabIndex        =   5
      Top             =   2445
      Width           =   5760
   End
   Begin VB.TextBox RegCode 
      Height          =   285
      Left            =   1110
      TabIndex        =   4
      Top             =   1065
      Width           =   3825
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Authorise"
      Default         =   -1  'True
      Height          =   495
      Left            =   2295
      TabIndex        =   0
      Top             =   1545
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   150
      X2              =   5910
      Y1              =   870
      Y2              =   870
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   5910
      Y1              =   885
      Y2              =   885
   End
   Begin VB.Label KeyCode 
      Caption         =   "#KEYCODE#"
      Height          =   255
      Left            =   3285
      TabIndex        =   3
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Registration Key Code:"
      Height          =   255
      Left            =   1485
      TabIndex        =   2
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "This copy of LanMan has not been authorized!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1035
      TabIndex        =   1
      Top             =   120
      Width           =   4005
   End
End
Attribute VB_Name = "Registration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim SerialCode
SerialCode = Hex(KeyCode.Caption)
If SerialCode = RegCode.Text Then
   SaveSetting "Lanman", "Settings", "Reg", SerialCode
Else
   End
End If
End Sub

Private Sub Form_Load()
File1.Path = "C:\Windows"
KeyCode.Caption = Asc((File1.ListCount * 2)) & Asc((File1.ListCount * 53))
End Sub

