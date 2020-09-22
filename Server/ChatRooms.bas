Attribute VB_Name = "ChatRooms"
Public RoomNames() As String
Public RoomMembers() As String
Public RoomPass() As String
Public BackupVar1() As String
Public BackupVar2() As String
Public BackupVar3() As String
Public Sub ExtendRoomArray()
On Error Resume Next
Dim BackupSize
Dim VarSize
BackupSize = 0
BackupSize = Int(UBound(RoomNames))
ReDim BackupVar1(0 To BackupSize)
ReDim BackupVar2(0 To BackupSize)
ReDim BackupVar3(0 To BackupSize)
For x = 0 To UBound(RoomNames)
BackupVar1(x) = RoomNames(x)
BackupVar2(x) = RoomMembers(x)
BackupVar3(x) = RoomPass(x)
Next x
VarSize = Int(UBound(RoomNames) + 1)
ReDim RoomNames(0 To VarSize)
ReDim RoomMembers(0 To VarSize)
ReDim RoomPass(0 To VarSize)
For x = 0 To UBound(RoomNames)
RoomNames(x) = BackupVar1(x)
RoomMembers(x) = BackupVar2(x)
RoomPass(x) = BackupVar3(x)
Next x
End Sub

Public Sub CreateRoom(Roomname As String, RoomType As String, Optional wsIndex As Integer, Optional Password As String)
ExtendRoomArray
Dim ArrayPosition
Dim sRoomName As String
ArrayPosition = UBound(RoomNames)
'// First see if the room already exists
For x = 0 To UBound(RoomNames)
If LCase(Roomname) = LCase(RoomNames(x)) Then
   If wsIndex <> "-1" Then
      SendClientCommand "RoomExists", wsIndex
      Exit Sub
   End If
End If
Next x
If RoomType = "Public" Then sRoomName = "0" & Chr(1) & Roomname
If RoomType = "Private" Then sRoomName = "1" & Chr(1) & Roomname
RoomNames(ArrayPosition) = sRoomName
If Password = "" Then Password = "<null>"
RoomPass(ArrayPosition) = Password
If MainForm.Option1.Value = True Then
   If RoomType = "Private" Then Exit Sub
End If
If MainForm.Option2.Value = True Then
   If RoomType = "Public" Then Exit Sub
End If
Set h = MainForm.Roomlist.ListItems.Add(, , Roomname)
h.SubItems(1) = GetMemberCount(Roomname)
h.SubItems(2) = RoomType
If wsIndex <> "-1" Then
   UserJoin Roomname, GetUsernameFromWS(wsIndex), wsIndex, Password
   Exit Sub
End If
End Sub

Public Sub RefreshRoomCount(Roomname As String)
For x = 1 To MainForm.Roomlist.ListItems.Count
If Roomname = MainForm.Roomlist.ListItems(x) Then
   MainForm.Roomlist.ListItems(x).SubItems(1) = GetMemberCount(Roomname)
   Exit Sub
End If
Next x
End Sub

Public Sub UserJoin(Roomname As String, Username As String, wsIndex As Integer, Password As String)
Dim RoomPos As Integer
RoomPos = LocateRoomInArray(Roomname)
If RoomPos = -1 Then
   SendClientCommand "BadRoom" & Chr(1) & Roomname, wsIndex
   Exit Sub
End If
If IsUserInRoom(Roomname, Username) = True Then Exit Sub

'// Check password
If LCase(RoomPass(RoomPos)) <> LCase(Password) Then
   If Left(RoomNames(RoomPos), 1) = 0 Then GoTo SkipCheck
   SendClientCommand "BadRoomPass" & Chr(1) & Roomname, wsIndex
   Exit Sub
End If
SkipCheck:
If RoomMembers(RoomPos) = "" Then
   RoomMembers(RoomPos) = Username
Else
   RoomMembers(RoomPos) = RoomMembers(RoomPos) & Chr(1) & Username
End If
SendClientCommand "RoomJoin" & Chr(1) & Roomname, wsIndex
RefreshRoomCount Roomname
End Sub

Public Function LocateUserInArray(Username As String) As Integer
For x = 0 To UBound(RoomMembers)
If LCase(RoomMembers(x)) = LCase(Username) Then
   LocateUserInArray = x
   Exit Function
End If
Next x
LocateUserInArray = -1
End Function

Public Function LocateRoomInArray(Roomname As String) As Integer
For x = 0 To UBound(RoomNames)
If Right(LCase(RoomNames(x)), Len(RoomNames(x)) - 2) = LCase(Roomname) Then
   LocateRoomInArray = x
   Exit Function
End If
Next x
LocateRoomInArray = -1
End Function

Public Function GetMemberCount(Roomname As String) As Integer
On Error Resume Next
Dim RawData As String
Dim tempdata As String
Dim EntryCount As Integer
EntryCount = 0
RawData = RoomMembers(LocateRoomInArray(Roomname))
If RawData = "" Then GetMemberCount = 0: Exit Function
Do Until RawData = ""
If InStr(1, RawData, Chr(1), vbTextCompare) <> 0 Then
    tempdata = Left(RawData, InStr(1, RawData, Chr(1), vbTextCompare) - 1)
    RawData = Right(RawData, Len(RawData) - Len(tempdata) - 1)
    EntryCount = EntryCount + 1
Else
    EntryCount = EntryCount + 1
    GoTo skiploop
End If
Loop
skiploop:
GetMemberCount = EntryCount
End Function

Public Function AssembleMemberList(Roomname As String) As String
On Error Resume Next
AssembleMemberList = RoomMembers(LocateRoomInArray(Roomname))
End Function

Public Sub PostRoomData(Roomname As String, CommandString As String, PostData As String)
Dim RawData As String
Dim tempdata As String
RawData = RoomMembers(LocateRoomInArray(Roomname))
Do Until RawData = ""
If InStr(1, RawData, Chr(1), vbTextCompare) <> 0 Then
   tempdata = Left(RawData, InStr(1, RawData, Chr(1), vbTextCompare) - 1)
   RawData = Right(RawData, Len(RawData) - Len(tempdata) - 1)
   DoEvents
   SendClientCommand CommandString & Chr(1) & Roomname & Chr(1) & PostData, GetWSFromUsername(tempdata)
Else
   tempdata = RawData
   DoEvents
   SendClientCommand CommandString & Chr(1) & Roomname & Chr(1) & PostData, GetWSFromUsername(tempdata)
   Exit Sub
End If
Loop
End Sub

Public Sub RemoveFromRoom(Roomname As String, Username As String)
    '// This procedure will remove a user from the
    '// specified chat-room.
    Dim sMessageData As String
    Dim sTemp As String
    Dim TempBuffer(0 To 500) As String
    Dim CounterBuffer As Integer
    
    ' First seperate and copy all users in the
    ' roommember array into another buffer
    CounterBuffer = 0
    sMessageData = RoomMembers(LocateRoomInArray(Roomname))
    Do Until sMessageData = ""
    If InStr(1, sMessageData, Chr(1), vbTextCompare) <> 0 Then
        sTemp = Left(sMessageData, InStr(1, sMessageData, Chr(1), vbTextCompare) - 1)
        sMessageData = Right(sMessageData, Len(sMessageData) - Len(sTemp) - 1)
        TempBuffer(CounterBuffer) = sTemp
    Else
        sTemp = sMessageData
        TempBuffer(CounterBuffer) = sTemp
        sMessageData = ""
        GoTo skipnext
    End If
    CounterBuffer = CounterBuffer + 1
    Loop
    
skipnext:
    ' Now search through all the seperated entries
    ' and take out the specified user
    For x = 0 To UBound(TempBuffer)
    If TempBuffer(x) = "" Then GoTo SkipEntry
    Dim rPacket As String
    If TempBuffer(x) <> Username Then
        If rPacket = "" Then
            rPacket = TempBuffer(x)
        Else
            rPacket = rPacket & Chr(1) & TempBuffer(x)
        End If
    End If
SkipEntry:
    Next x
    
    ' Finally copy all contents of the temp buffer
    ' back into the roommember array
    RoomMembers(LocateRoomInArray(Roomname)) = rPacket
    
    ' Now refresh the room details and post to all
    ' users in the chatroom a message
    PostRoomData Roomname, "UserRoomRemove", Username
    DoEvents
    RefreshRoomCount Roomname
    Exit Sub
End Sub

Public Function IsUserInRoom(Roomname As String, Username As String) As Boolean
    '// This function will check to see if a user
    '// is present within a chatroom.
    '// It will return either true or false.
    Dim sMessageData As String
    Dim sTemp As String
    Dim TempBuffer() As String
    Dim CounterBuffer As Integer
    CounterBuffer = 0
    If GetMemberCount(Roomname) = 0 Then
        IsUserInRoom = False
        Exit Function
    End If
    sMessageData = RoomMembers(LocateRoomInArray(Roomname))
    Do Until sMessageData = ""
    ReDim TempBuffer(CounterBuffer)
    If InStr(1, sMessageData, Chr(1), vbTextCompare) <> 0 Then
        sTemp = Left(sMessageData, InStr(1, sMessageData, Chr(1), vbTextCompare) - 1)
        sMessageData = Right(sMessageData, Len(sMessageData) - Len(sTemp) - 1)
        TempBuffer(CounterBuffer) = sTemp
    Else
        sTemp = sMessageData
        TempBuffer(CounterBuffer) = sTemp
        GoTo skipnext
    End If
    CounterBuffer = CounterBuffer + 1
    Loop
skipnext:
    For x = 0 To UBound(TempBuffer)
    If LCase(TempBuffer(x)) = LCase(Username) Then IsUserInRoom = True: Exit Function
    Next x
    IsUserInRoom = False
End Function

Public Sub RefreshRoomList(Filter As String)
If Filter = "Public" Then
   ShowPublicRooms
   Exit Sub
End If
If Filter = "Private" Then
   ShowPrivateRooms
   Exit Sub
End If
End Sub

Private Sub ShowPublicRooms()
Dim TempBuffer As String
MainForm.Roomlist.ListItems.Clear
For x = 0 To UBound(RoomNames)
TempBuffer = RoomNames(x)
If Left(TempBuffer, 1) = 0 Then
   Set h = MainForm.Roomlist.ListItems.Add(, , Right(TempBuffer, Len(TempBuffer) - 2))
   h.SubItems(1) = 0
   h.SubItems(2) = "Public"
   RefreshRoomCount Right(TempBuffer, Len(TempBuffer) - 2)
End If
Next x
End Sub

Private Sub ShowPrivateRooms()
Dim TempBuffer As String
MainForm.Roomlist.ListItems.Clear
For x = 0 To UBound(RoomNames)
TempBuffer = RoomNames(x)
If Left(TempBuffer, 1) = 1 Then
   Set h = MainForm.Roomlist.ListItems.Add(, , Right(TempBuffer, Len(TempBuffer) - 2))
   h.SubItems(1) = 0
   h.SubItems(2) = "Private"
   RefreshRoomCount Right(TempBuffer, Len(TempBuffer) - 2)
End If
Next x
End Sub

Public Sub ShowAllRooms()
Dim TempBuffer As String
Dim RoomType As String
MainForm.Roomlist.ListItems.Clear
For x = 0 To UBound(RoomNames)
TempBuffer = RoomNames(x)
If Left(TempBuffer, 1) = 0 Then RoomType = "Public"
If Left(TempBuffer, 1) = 1 Then RoomType = "Private"
Set h = MainForm.Roomlist.ListItems.Add(, , Right(TempBuffer, Len(TempBuffer) - 2))
h.SubItems(1) = 0
h.SubItems(2) = RoomType
RefreshRoomCount Right(TempBuffer, Len(TempBuffer) - 2)
Next x
End Sub

