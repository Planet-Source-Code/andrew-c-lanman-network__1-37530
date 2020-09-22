Attribute VB_Name = "Functions"
Public Declare Function SetCursorPos Lib "USER32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetClassName Lib "USER32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowRect Lib "USER32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
    End Type
Private What As RECT
Public Sub LoadUserList()
On Error GoTo 1
MainForm.UserList.Nodes.Clear
Open "user.ini" For Input As #1
Do While Not EOF(1)
Input #1, temp
MainForm.UserList.Nodes.Add , , , temp
Loop
1 Close #1
End Sub

Public Sub RemoveUser(Username As String)
On Error GoTo 1
Open "user.ini" For Input As #1
Open "user.tmp" For Output As #2
Close #2
Open "user.tmp" For Append As #2
Do While Not EOF(1)
Input #1, temp
If temp <> Username Then Print #2, temp
Loop
Close #1
Close #2
Kill "user.ini"
Open "user.tmp" For Input As #1
Open "user.ini" For Append As #2
Do While Not EOF(1)
Input #1, temp
Print #2, temp
Loop
Close #1
Close #2
Kill "user.tmp"
LoadUserList
1 End Sub

Public Function ConvertUserToIP(Username As String)
Dim RegUser
RegUser = GetSetting("Lanman", "Users", Username)
On Error GoTo 1
Open "user.ini" For Input As #1
Do While Not EOF(1)
Input #1, temp
If temp = Username Then
    ConvertUserToIP = RegUser
    Close #1
    Exit Function
End If
Loop
1     Close #1
End Function

Public Function AddNewUser(Username As String, IPAddress As String)
For X = 1 To MainForm.UserList.Nodes.Count
If Username = MainForm.UserList.Nodes(X).Text Then
   Exit Function
End If
Next X
Open ("User.ini") For Append As #1
Print #1, Username
SaveSetting "Lanman", "Users", Username, IPAddress
Close #1
LoadUserList
End Function

Public Function DoTime()
Dim TimeVar
Dim LeftSec
Dim RightSec
TimeVar = FormatDateTime(Time$, vbLongTime)
DoTime = TimeVar
End Function
Public Function GetStartMenuHeight() As Long
Dim HeightofStartMenu
Dim i, X, Y, z$
    For i = 1 To 999 '// The start menu never uses a HWND higher than 1000
        z$ = Space$(128)
        Y = GetClassName(i, z$, 128)
        X = Left$(z$, Y)
        If LCase(X) = "shell_traywnd" Then
            GoTo JumpOut:
        End If
    Next i
JumpOut:
    GetWindowRect i, What
    HeightofStartMenu = What.Top * 15
GetStartMenuHeight = HeightofStartMenu
End Function

Public Function PopupWindow(Text As String, Title As String)
Popup.Title = Title
Popup.Message = Text
End Function

Public Function SendChatText(Text As String)
On Error Resume Next
Dim strData As String
strData = Settings.txtusername & ":  " & ChatWnd.Message.Text
MainForm.Chat.SendData strData
ChatWnd.Conversation.Text = ChatWnd.Conversation.Text & vbCrLf & strData
ChatWnd.Conversation.SetFocus
ChatWnd.Conversation.SelStart = Len(ChatWnd.Conversation.Text)
ChatWnd.Message.Text = ""
ChatWnd.Message.SetFocus
End Function
Public Sub LogOn()
Dim strData As String
Dim broadcastbuffer
MainForm.LogOnMessage.Visible = False
MainForm.LogOnMessage.Enabled = False
MainForm.UserList.Enabled = True
LoadUserList
MainForm.Chat.Close
MainForm.Chat.Listen
strData = Settings.txtusername.Text
MainForm.OnlineStrip1.Visible = True
MainForm.OnlineStrip2.Visible = True
MainForm.Image1.Enabled = True
MainForm.Image2.Enabled = True
Resource.onlinestatus.Caption = "true"
broadcastbuffer = MainForm.Broadcast.Write(strData, Len(strData))
End Sub

Public Sub LogOff()
MainForm.LogOnMessage.Visible = True
MainForm.LogOnMessage.Enabled = True
MainForm.UserList.Nodes.Clear
MainForm.Chat.Close
MainForm.OnlineStrip1.Visible = False
MainForm.OnlineStrip2.Visible = False
MainForm.Image1.Enabled = False
MainForm.Image2.Enabled = False
Resource.onlinestatus.Caption = "false"
End Sub

Public Sub LoadEverything()
'// This procedure sets all properties/settings within
'// the program. It also prepares the sockets:
'// Place any pre-load code here!

Dim TopPos
Dim LeftPos
TopPos = GetSetting("Lanman", "Settings", "YPos")
LeftPos = GetSetting("Lanman", "Settings", "XPos")
If TopPos = "" Or LeftPos = "" Then
    MainForm.Top = 0
    MainForm.Left = 0
    GoTo endofMove
End If
MainForm.Top = TopPos
MainForm.Left = LeftPos
endofMove:
Dim Username
Username = GetSetting("Lanman", "Settings", "Nickname")
If Username = "" Then
   Settings.FirstTimeMessage.Visible = True
   Settings.Show
   'Timer1.Enabled = False
   Resource.SettingsFirst.Caption = "yes"
   Settings.ZOrder (0)
   Exit Sub
End If
ShowIcon MainForm
Settings.txtusername.Text = Username
Settings.AutoStart.Value = GetSetting("Lanman", "Settings", "autostart")
AddUser.sckBroadcast.LocalPort = 2000
AddUser.sckBroadcast.RemotePort = 2000
AddUser.sckBroadcast.AddressFamily = AF_INET
AddUser.sckBroadcast.Protocol = IPPROTO_UDP
AddUser.sckBroadcast.SocketType = SOCK_DGRAM
AddUser.sckBroadcast.Broadcast = True
AddUser.sckBroadcast.Binary = False
AddUser.sckBroadcast.Blocking = False
AddUser.sckBroadcast.Action = SOCKET_OPEN

MainForm.Broadcast.LocalPort = 2001
MainForm.Broadcast.RemotePort = 2001
MainForm.Broadcast.AddressFamily = AF_INET
MainForm.Broadcast.Protocol = IPPROTO_UDP
MainForm.Broadcast.SocketType = SOCK_DGRAM
MainForm.Broadcast.Broadcast = True
MainForm.Broadcast.Binary = False
MainForm.Broadcast.Blocking = False
MainForm.Broadcast.Action = SOCKET_OPEN

MainForm.Messenger.LocalPort = 2002
MainForm.Messenger.RemotePort = 2002
MainForm.Messenger.AddressFamily = AF_INET
MainForm.Messenger.Protocol = IPPROTO_UDP
MainForm.Messenger.SocketType = SOCK_DGRAM
MainForm.Messenger.Broadcast = True
MainForm.Messenger.Binary = False
MainForm.Messenger.Blocking = False
MainForm.Messenger.Action = SOCKET_OPEN
End Sub

Sub Main()
'StartupCheck
CheckSerialKey
End Sub

Public Function TerminateForms()
Unload About
Unload AddUser
Unload ChatWnd
Unload MainForm
Unload Popup
Unload Resource
Unload Settings
Unload Startup
End Function

Public Sub StartupCheck()
If GetSetting("Lanman", "Settings", "ShowStartup") = "0" Then
    LoadEverything
    MainForm.Show
Else
    tartup.Show
End If
End Sub
