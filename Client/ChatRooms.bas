Attribute VB_Name = "ChatRooms"
Public WindowBuffers(0 To 20) As ChatWnd 'limit of 20 windows
Public Checkers(0 To 20) As Boolean
Public RoomMatches(0 To 20) As String

Public Function FindFreeHandle() As Integer
For X = 0 To 20
If Checkers(X) = False Then
   FindFreeHandle = X
   Exit Function
End If
Next X
FindFreeHandle = -1
End Function

Public Function GetWndhandle(Roomname As String) As Integer
For X = 0 To 20
If LCase(Roomname) = LCase(RoomMatches(X)) Then
   GetWndhandle = X
End If
Next X
End Function
