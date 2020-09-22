Attribute VB_Name = "SystrayAPI"
'// System Tray Functions Module
'Option Explicit
Private FormHandle As Long
Private mvarbRunningInTray As Boolean
Private SysIcon As NOTIFYICONDATA
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Property Let bRunningInTray(ByVal vData As Boolean)
    mvarbRunningInTray = vData
End Property

Property Get bRunningInTray() As Boolean
    bRunningInTray = mvarbRunningInTray
End Property

Public Sub ShowIcon(ByRef Systrayform As Form)
    SysIcon.cbSize = Len(SysIcon)
    SysIcon.hwnd = Systrayform.hwnd
    SysIcon.uId = vbNull
    SysIcon.uFlags = 7
    SysIcon.ucallbackMessage = 512
    SysIcon.hIcon = Systrayform.Icon
    SysIcon.szTip = Systrayform.Caption + Chr(0)
    Shell_NotifyIcon 0, SysIcon
    mvarbRunningInTray = True
End Sub

Public Sub RemoveIcon(Systrayform As Form)
    SysIcon.cbSize = Len(SysIcon)
    SysIcon.hwnd = Systrayform.hwnd
    SysIcon.uId = vbNull
    SysIcon.uFlags = 7
    SysIcon.ucallbackMessage = vbNull
    SysIcon.hIcon = Systrayform.Icon
    SysIcon.szTip = Chr(0)
    Shell_NotifyIcon 2, SysIcon
    If Systrayform.Visible = False Then Systrayform.Show    'Incase user can't see form
    mvarbRunningInTray = False
End Sub

Public Sub ChangeIcon(Systrayform As Form, picNewIcon As PictureBox)
    If mvarbRunningInTray = True Then
        SysIcon.cbSize = Len(SysIcon)
        SysIcon.hwnd = Systrayform.hwnd
        SysIcon.hIcon = picNewIcon.Picture
        Shell_NotifyIcon 1, SysIcon
    End If
End Sub

Public Sub ChangeToolTip(Systrayform As Form, strNewTip As String)
    If mvarbRunningInTray = True Then
        SysIcon.cbSize = Len(SysIcon)
        SysIcon.hwnd = Systrayform.hwnd
        SysIcon.szTip = strNewTip & Chr(0)
        Shell_NotifyIcon 1, SysIcon
    End If
End Sub

