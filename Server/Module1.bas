Attribute VB_Name = "MainFunctions"
Public Sub StartServer()
MainForm.wsServer.Listen
CreateRoom "Lanman Public Lounge", "Public"
CreateRoom "Help desk", "Public"
CreateRoom "Engine room", "Private", -1, "redhat"
End Sub

Public Sub StopServer()
MainForm.wsServer.Close
For x = 0 To MainForm.wsClient.Count - 1
MainForm.wsClient(x).Close
Next x
End Sub

Public Function LocateUserFromWS(wsString As Integer)
For x = 1 To MainForm.UserList.ListItems.Count
If wsString = MainForm.UserList.ListItems(x).SubItems(1) Then
    LocateUserFromWS = x
    Exit Function
End If
Next x
End Function

Public Function GetUsernameFromWS(wsIndex As Integer)
On Error Resume Next
GetUsernameFromWS = MainForm.UserList.ListItems(LocateUserFromWS(wsIndex)).Text
End Function

Public Function GetWSFromUsername(Username As String)
For x = 1 To MainForm.UserList.ListItems.Count
If LCase(Username) = LCase(MainForm.UserList.ListItems(x).Text) Then
   GetWSFromUsername = MainForm.UserList.ListItems(x).SubItems(1)
   Exit Function
End If
Next x
End Function

Public Function FileSizeConv(InputText As String) As String
'// DOTO: Code function to convert and format filesize

'FileSizeConv = InputText & " Bytes"
End Function

Public Sub CheckPaths()
On Error Resume Next
MkDir App.Path & "\Accounts"
MkDir App.Path & "\Messages"
MkDir App.Path & "\Messages\Public"
End Sub
