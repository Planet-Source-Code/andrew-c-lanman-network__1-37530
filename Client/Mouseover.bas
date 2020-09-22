Attribute VB_Name = "Mouseover"
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Function GetHandle() As Long
Dim CursorPOS As POINTAPI
GetCursorPos CursorPOS
GetHandle = WindowFromPoint(CursorPOS.X, CursorPOS.Y)
End Function
