Attribute VB_Name = "Functions"
'// Main Programs Module
Public Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
'Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private What As RECT
'Global lRegion As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public TextString As String
Public TextPOS As Integer
Public Spaces As Integer
Public PackPOS As Integer
Public Packing As Boolean
Public PackString As String
Public InitText As String
Public AwaitingChat As Boolean
Public Accepted As Boolean
Public DenyAll As Boolean
Public BlinkCount As Integer
Public BlinkCountCheck As Boolean
Public Loaded As Boolean
Public UserStatus As String
Public StartupMain As Boolean
Public DisableFunctions As Boolean
Public IPValue As String
Public Proceed As Boolean

Sub Main()
    Startup.Initialize
    'LoadEverything
End Sub

Public Sub CenterForm(Who As Form)
    On Local Error Resume Next
    Who.Left = (Screen.Width - Who.Width) / 2
    Who.Top = (Screen.Height - Who.Height) / 2
End Sub

Public Sub DragForm(frm As Form)
    On Local Error Resume Next
    ' Snap to coding
    ' Snap code has been removed in this release.
End Sub

Public Function DoTime()
    Dim TimeVar
    Dim LeftSec
    Dim RightSec
    TimeVar = FormatDateTime(Time$, vbLongTime)
    DoTime = TimeVar
End Function
Public Function GetStartMenuHeightA() As Long
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
    GetStartMenuHeightA = (Screen.Height - HeightofStartMenu)
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
    'Popup.Title = Title
    'Popup.Message = Text
    TextString = Text
    InitText = Text
    Spaces = 0
    Packing = False
    PackPOS = 0
    TextPOS = 1
    BlinkCount = 0
'    MainForm.TextScroll.Alignment = 1
    MainForm.TextScroll.Caption = Text
    MainForm.ScrollingText.Enabled = True
End Function

Public Sub LogOn()
    Dim strData As String
    Dim broadcastbuffer
    'MainForm.LoginFrame.Visible = False
    'MainForm.ButtonBar.Visible = True
    'LoadUserList
    UserStatus = "Available"
    PopupWindow "Lanman is online", "System Message"
    Resource.onlinestatus.Caption = "true"
    MainForm.TitleText.Caption = "Lanman Network - <" & frmLogin.Username.Text & ">"
End Sub

Public Sub LogOff()
    On Error Resume Next
    'MainForm.LoginFrame.Visible = True
    Resource.onlinestatus.Caption = "false"
    'MainForm.ButtonBar.Visible = False
    UserStatus = "Offline"
    PopupWindow "Lanman is offline", "System Message"
    SendServerCommand "LogOut"
    DoEvents
    MainForm.Server.Close
End Sub

Public Sub LoadEverything()
    On Error Resume Next
    LogOff
    'InitNetCards
    InitSockets
    InitForms
    InitSettings
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

Public Sub LoadSkinData(SkinString As String)
On Error GoTo SkinError
SkinRes.LoginPicture.Picture = LoadPicture(App.Path & "\Skinz\" & SkinString & "\" & GetINIVal("SkinFiles", "00", App.Path & "\Skinz\" & SkinString & "\Skin.dat"))
SkinRes.MainWindow.Picture = LoadPicture(App.Path & "\Skinz\" & SkinString & "\" & GetINIVal("SkinFiles", "01", App.Path & "\Skinz\" & SkinString & "\Skin.dat"))
SkinRes.Button0.Picture = LoadPicture(App.Path & "\Skinz\" & SkinString & "\" & GetINIVal("SkinFiles", "02", App.Path & "\Skinz\" & SkinString & "\Skin.dat"))
SkinRes.Button1.Picture = LoadPicture(App.Path & "\Skinz\" & SkinString & "\" & GetINIVal("SkinFiles", "03", App.Path & "\Skinz\" & SkinString & "\Skin.dat"))
SkinRes.Button2.Picture = LoadPicture(App.Path & "\Skinz\" & SkinString & "\" & GetINIVal("SkinFiles", "04", App.Path & "\Skinz\" & SkinString & "\Skin.dat"))
SkinRes.Button3.Picture = LoadPicture(App.Path & "\Skinz\" & SkinString & "\" & GetINIVal("SkinFiles", "05", App.Path & "\Skinz\" & SkinString & "\Skin.dat"))
SkinRes.Button4.Picture = LoadPicture(App.Path & "\Skinz\" & SkinString & "\" & GetINIVal("SkinFiles", "06", App.Path & "\Skinz\" & SkinString & "\Skin.dat"))
MainForm.LoginFrame.Picture = SkinRes.LoginPicture.Picture
MainForm.Image3.Picture = SkinRes.MainWindow.Picture
MainForm.ButtonBar.Picture = SkinRes.Button0.Picture
SkinRes.TopBarLeft0.Picture = LoadPicture(App.Path & "\Skinz\" & SkinString & "\" & GetINIVal("SkinFiles", "08", App.Path & "\Skinz\" & SkinString & "\Skin.dat"))
SkinRes.TopBarLeft1.Picture = LoadPicture(App.Path & "\Skinz\" & SkinString & "\" & GetINIVal("SkinFiles", "09", App.Path & "\Skinz\" & SkinString & "\Skin.dat"))
SkinRes.TopBarTile0.Picture = LoadPicture(App.Path & "\Skinz\" & SkinString & "\" & GetINIVal("SkinFiles", "10", App.Path & "\Skinz\" & SkinString & "\Skin.dat"))
SkinRes.TopbarTile1.Picture = LoadPicture(App.Path & "\Skinz\" & SkinString & "\" & GetINIVal("SkinFiles", "11", App.Path & "\Skinz\" & SkinString & "\Skin.dat"))
SkinRes.TopBarRight0.Picture = LoadPicture(App.Path & "\Skinz\" & SkinString & "\" & GetINIVal("SkinFiles", "12", App.Path & "\Skinz\" & SkinString & "\Skin.dat"))
SkinRes.TopbarRight1.Picture = LoadPicture(App.Path & "\Skinz\" & SkinString & "\" & GetINIVal("SkinFiles", "13", App.Path & "\Skinz\" & SkinString & "\Skin.dat"))
SkinRes.imgkrnbl.Picture = LoadPicture(App.Path & "\Skinz\" & SkinString & "\" & GetINIVal("SkinFiles", "14", App.Path & "\Skinz\" & SkinString & "\Skin.dat"))
SkinRes.imgkrnbr.Picture = LoadPicture(App.Path & "\Skinz\" & SkinString & "\" & GetINIVal("SkinFiles", "15", App.Path & "\Skinz\" & SkinString & "\Skin.dat"))
SkinRes.imgkrntl.Picture = LoadPicture(App.Path & "\Skinz\" & SkinString & "\" & GetINIVal("SkinFiles", "16", App.Path & "\Skinz\" & SkinString & "\Skin.dat"))
SkinRes.imgkrntr.Picture = LoadPicture(App.Path & "\Skinz\" & SkinString & "\" & GetINIVal("SkinFiles", "17", App.Path & "\Skinz\" & SkinString & "\Skin.dat"))
SkinRes.imgkrnrv.Picture = LoadPicture(App.Path & "\Skinz\" & SkinString & "\" & GetINIVal("SkinFiles", "18", App.Path & "\Skinz\" & SkinString & "\Skin.dat"))
SkinRes.imgkrnlv.Picture = LoadPicture(App.Path & "\Skinz\" & SkinString & "\" & GetINIVal("SkinFiles", "19", App.Path & "\Skinz\" & SkinString & "\Skin.dat"))
SkinRes.imgkrnwnd.Picture = LoadPicture(App.Path & "\Skinz\" & SkinString & "\" & GetINIVal("SkinFiles", "20", App.Path & "\Skinz\" & SkinString & "\Skin.dat"))
SkinRes.imgkrnbh.Picture = LoadPicture(App.Path & "\Skinz\" & SkinString & "\" & GetINIVal("SkinFiles", "21", App.Path & "\Skinz\" & SkinString & "\Skin.dat"))
SkinRes.imgkrnth.Picture = LoadPicture(App.Path & "\Skinz\" & SkinString & "\" & GetINIVal("SkinFiles", "22", App.Path & "\Skinz\" & SkinString & "\Skin.dat"))
Dim bgr As String
Dim bgg As String
Dim bgb As String
bgr = GetINIVal("SkinFiles", "BGR", App.Path & "\Skinz\" & SkinString & "\Skin.dat")
bgg = GetINIVal("SkinFiles", "BGG", App.Path & "\Skinz\" & SkinString & "\Skin.dat")
bgb = GetINIVal("SkinFiles", "BGB", App.Path & "\Skinz\" & SkinString & "\Skin.dat")
SkinRes.BGColor.BackColor = RGB(bgr, bgg, bgb)
SkinAll
Exit Sub
SkinError:
MsgBox "Skin Error:" & vbCrLf & "The " & SkinString & " skin is damaged or incomplete", vbCritical, "Cannot load skin data"
End
End Sub

Public Sub SkinWindow(Window As Form)
Window.TopBarLeft.Picture = SkinRes.TopBarLeft0.Picture
Window.TopBarTile.Picture = SkinRes.TopBarTile0.Picture
Window.TopBarRight.Picture = SkinRes.TopBarRight0.Picture
Window.Picture1.Picture = SkinRes.imgkrntl.Picture
Window.Image1.Picture = SkinRes.imgkrnth.Picture
Window.Picture2.Picture = SkinRes.imgkrntr.Picture
Window.Image2.Picture = SkinRes.imgkrnlv.Picture
Window.Image5.Picture = SkinRes.imgkrnrv.Picture
Window.Picture3.Picture = SkinRes.imgkrnbl.Picture
Window.Picture4.Picture = SkinRes.imgkrnbr.Picture
Window.Image6.Picture = SkinRes.imgkrnbh.Picture
Window.BackColor = SkinRes.BGColor.BackColor
End Sub

Public Sub SkinAll()
SkinWindow ChatWnd
SkinWindow frmConsole
SkinWindow MessageBoard
SkinWindow RoomPass
SkinWindow RoomList
End Sub

Public Function CheckDisabled() As Boolean
If DisableFunctions = True Then
   MsgBox "No network adapters found on system." & vbCrLf & "This feature is not available.", vbExclamation, "Networking Failure"
   CheckDisabled = True
End If
End Function

Public Sub NeedRestart()
MsgBox "Lanman must restart for the changes to take effect.", vbExclamation, "Application Settings"
MainForm.Hide
Startup.Initialize
End Sub

Public Sub InitSockets()
Proceed = False
End Sub

Public Sub InitForms()
    '// Load forms and place in systray
    ShowIcon MainForm
    Load AddUser
    Load MainForm
    Load AppSettings
    DoEvents
End Sub

Public Sub InitSettings()
    '// Initialize all program settings
    Proceed = True
    Dim TopPos
    Dim LeftPos
    Dim Username
    TopPos = GetKey("Lanman", "Settings", "YPos")
    LeftPos = GetKey("Lanman", "Settings", "XPos")
        If TopPos = "" Or LeftPos = "" Then
        MainForm.Top = 0
        MainForm.Left = 0
        GoTo endofMove
        End If
    MainForm.Top = TopPos
    MainForm.Left = LeftPos
endofMove:
    AppSettings.AutoStart.Value = GetKey("Lanman", "Settings", "autostart", 0)
    MainForm.LoginFrame.Top = 240
    MainForm.LoginFrame.Left = -15
    MainForm.Height = 2190
    Dim LogonSetting
    Dim StartMinValue
    StartMinValue = GetKey("LanMan", "Settings", "StartMin", 0)
    AppSettings.Startmin.Value = StartMinValue
    LogonSetting = GetKey("Lanman", "Settings", "AutoLogon", 0)
    AppSettings.AutoLogon.Value = LogonSetting
    frmLogin.Check1.Value = LogonSetting
endsubfunc:
Proceed = False
End Sub

Public Sub ErrorPopup(SString As String)
MsgBox "Unable to complete request:" & vcrlf & vbCrLf & SString, vbInformation, "Operation failed!"
End Sub
