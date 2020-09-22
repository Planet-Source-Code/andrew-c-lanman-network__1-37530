Attribute VB_Name = "NetConnect"
Public Function FindFreeSock()
'// First search through all existing sockets
For x = 0 To MainForm.wsClient.Count - 1
If MainForm.wsClient(x).State <> 7 Then
    '// Found a socket!
    FindFreeSock = x
    MainForm.wsClient(x).Close
    Exit Function
End If
Next x
'// Create a new socket
Load MainForm.wsClient(MainForm.wsClient.Count)
FindFreeSock = MainForm.wsClient.Count - 1
End Function

Public Sub SendClientCommand(StringData As String, wsIndex As Integer)
On Error Resume Next
MainForm.wsClient(wsIndex).SendData "#" & StringData
DoEvents
End Sub

Public Sub ProcessLoginDetails(RawData As String, wsIndex As Integer)
Dim Username As String
Dim Password As String
Username = Left(RawData, InStr(1, RawData, Chr(1), vbTextCompare) - 1)
Password = Right(RawData, Len(RawData) - Len(Username) - 1)

'// See if they are already logged on
For x = 1 To MainForm.UserList.ListItems.Count
If LCase(Username) = LCase(MainForm.UserList.ListItems(x)) Then
   SendClientCommand "AlreadyOnline", wsIndex
   Exit Sub
End If
Next x

MainForm.AccountList.Path = App.Path & "\Accounts"
MainForm.AccountList.Refresh
For x = 0 To MainForm.AccountList.ListCount - 1
If LCase(Username) = LCase(Left(MainForm.AccountList.List(x), (Len(MainForm.AccountList.List(x)) - 4))) Then
   '// Check account
   Open App.Path & "\Accounts\" & MainForm.AccountList.List(x) For Input As #1
   Dim sUsername As String
   Dim sPassword As String
   Input #1, sUsername
   Input #1, sPassword
   If LCase(sPassword) = LCase(Password) Then
      LoginUser sUsername, wsIndex
   Else
      SendClientCommand "BadPassword", wsIndex
      Close #1
      Exit Sub
   End If
   Close #1
   Exit Sub
End If
Next x
'// User account not found
SendClientCommand "NoAccount", wsIndex
End Sub

Public Sub CreateLoginAccount(RawData As String, wsIndex As Integer)
Dim Username As String
Dim Password As String
Username = Left(RawData, InStr(1, RawData, Chr(1), vbTextCompare) - 1)
Password = Right(RawData, Len(RawData) - Len(Username) - 1)
'// See if it already exists
MainForm.AccountList.Path = App.Path & "\Accounts"
MainForm.AccountList.Refresh
For x = 0 To MainForm.AccountList.ListCount - 1
If Username = Left(MainForm.AccountList.List(x), (Len(MainForm.AccountList.List(x)) - 4)) Then
   SendClientCommand "AccountExists", wsIndex
   Exit Sub
End If
Next x
Open App.Path & "\Accounts\" & Username & ".lcn" For Output As #1
Print #1, Username
Print #1, Password
Close #1
DoEvents
MkDir App.Path & "\Messages\" & Username
SendClientCommand "CreatedOK", wsIndex
Exit Sub
End Sub

Public Sub LoginUser(Username As String, wsIndex As Integer)
Set D = MainForm.UserList.ListItems.Add(, , Username)
D.SubItems(1) = wsIndex
SendClientCommand "LoginOK", wsIndex
End Sub
