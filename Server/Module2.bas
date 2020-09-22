Attribute VB_Name = "Protocol"
Dim MessageID As String
Dim RawData As String
Dim sMailbox As String
Public Sub ProcessCommand(sString As String, wsIndex As Integer)
On Error Resume Next
Select Case sString
'
Case "#RequestLogin"
    SendClientCommand "SendLogin", wsIndex
    Exit Sub
Case "#LogOut"

    Dim sdUsername As String
    sdUsername = MainForm.UserList.ListItems(LocateUserFromWS(wsIndex)).Text
    MainForm.UserList.ListItems.Remove LocateUserFromWS(wsIndex)
    Dim RemovalCount:
Parse:
    RemovalCount = 0
    For x = 1 To MainForm.FileList.ListItems.Count
    If MainForm.FileList.ListItems(x).SubItems(2) = sdUsername Then
       MainForm.FileList.ListItems.Remove x
       RemovalCount = RemovalCount + 1
       GoTo parse2
    End If
    Next x
parse2:
    If RemovalCount <> 0 Then GoTo Parse
    Exit Sub
Case "#DoneFileList"
    FServe.RefreshTotalNetHost
    Exit Sub
Case "#SearchProceed"
    DoEvents
    MainForm.UserList.ListItems(LocateUserFromWS(wsIndex)).SubItems(2) = "OK"
    DoEvents
    Exit Sub
Case "#GetRoomCount"
    SendClientCommand "SetRoomCount" & Chr(1) & MainForm.Roomlist.ListItems.Count, wsIndex
    Exit Sub
Case "#KillSearch"
    MainForm.UserList.ListItems(LocateUserFromWS(wsIndex)).SubItems(2) = "Kill"
    Exit Sub
End Select
Dim RawData As String
RawData = Right(sString, Len(sString) - (Len(Left(sString, InStr(1, sString, Chr(1), vbTextCompare) - 1)) + 1))
Select Case Left(sString, InStr(1, sString, Chr(1), vbTextCompare) - 1)
'
Case "#Login"
    ProcessLoginDetails Right(sString, Len(sString) - 7), wsIndex
    Exit Sub
Case "#CreateLogin"
    CreateLoginAccount Right(sString, Len(sString) - 13), wsIndex
    Exit Sub
Case "#PrivateMessage"
    ProcessPrivateMessage (Right(sString, Len(sString) - 16)), wsIndex
    Exit Sub
Case "#GetPrivateHeaders"
    Dim HeaderData As String
    sMailbox = RawData
    MainForm.TempFile.Path = App.Path & "\Messages\" & sMailbox
    MainForm.TempFile.Refresh
    HeaderData = MainForm.TempFile.List(0)
    For x = 1 To MainForm.TempFile.ListCount - 1
    HeaderData = HeaderData & Chr(1) & MainForm.TempFile.List(x)
    Next x
    If HeaderData = "" Then HeaderData = "<Null>"
    HeaderData = MainForm.TempFile.ListCount & Chr(1) & HeaderData
    SendClientCommand "PrivateHeaders" & Chr(1) & HeaderData, wsIndex
    DoEvents
    Exit Sub
Case "#GetPublicHeaders"
    sMailbox = RawData
    MainForm.TempFile.Path = App.Path & "\Messages"
    MainForm.TempFile.Refresh
    HeaderData = MainForm.TempFile.List(0)
    For x = 1 To MainForm.TempFile.ListCount - 1
    HeaderData = HeaderData & Chr(1) & MainForm.TempFile.List(x)
    Next x
    If HeaderData = "" Then HeaderData = "<Null>"
    HeaderData = MainForm.TempFile.ListCount & Chr(1) & HeaderData
    SendClientCommand "PublicHeaders" & Chr(1) & HeaderData, wsIndex
    DoEvents
    Exit Sub
Case "#GetMessageTitle"
    On Error Resume Next
    MessageID = Left(RawData, InStr(1, RawData, Chr(1), vbTextCompare) - 1)
    sMailbox = Right(RawData, Len(RawData) - Len(MessageID) - 1)
    Dim SentDate As String
    Dim SentTime As String
    Dim SentFrom As String
    Dim Subject
    Dim Priority
    Dim NullString As String
    Open App.Path & "\Messages\" & sMailbox & "\" & MessageID For Input As #1
    Input #1, SentDate
    Input #1, SentTime
    Input #1, Priority
    Input #1, NullString
    Input #1, SentFrom
    Input #1, Subject
    Close #1
    SentDate = Right(SentDate, Len(SentDate) - 12)
    SentTime = Right(SentTime, Len(SentTime) - 12)
    Priority = Right(Priority, Len(Priority) - 5)
    SentFrom = Right(SentFrom, Len(SentFrom) - 5)
    Subject = Right(Subject, Len(Subject) - 9)
    SendClientCommand "AddMessage" & Chr(1) & MessageID & Chr(1) & Priority & Chr(1) & SentDate & Chr(1) & SentTime & Chr(1) & SentFrom & Chr(1) & Subject, wsIndex
    Exit Sub
Case "#GetPublicTitle"
    On Error Resume Next
    MessageID = Left(RawData, InStr(1, RawData, Chr(1), vbTextCompare) - 1)
    sMailbox = Right(RawData, Len(RawData) - Len(MessageID) - 1)
    Open App.Path & "\Messages\" & MessageID For Input As #1
    Input #1, SentDate
    Input #1, SentTime
    Input #1, Priority
    Input #1, NullString
    Input #1, SentFrom
    Input #1, Subject
    Close #1
    SentDate = Right(SentDate, Len(SentDate) - 12)
    SentTime = Right(SentTime, Len(SentTime) - 12)
    Priority = Right(Priority, Len(Priority) - 5)
    SentFrom = Right(SentFrom, Len(SentFrom) - 5)
    Subject = Right(Subject, Len(Subject) - 9)
    SendClientCommand "AddMessage" & Chr(1) & MessageID & Chr(1) & Priority & Chr(1) & SentDate & Chr(1) & SentTime & Chr(1) & SentFrom & Chr(1) & Subject, wsIndex
    Exit Sub
Case "#GetMessage"
    Dim BoxType As String
    Dim FileData As String
    Dim tempdata As String
    Dim gMessageID As String
    gMessageID = Left(RawData, InStr(1, RawData, Chr(1), vbTextCompare) - 1)
    RawData = Right(RawData, Len(RawData) - Len(gMessageID) - 1)
    sMailbox = Left(RawData, InStr(1, RawData, Chr(1), vbTextCompare) - 1)
    RawData = Right(RawData, Len(RawData) - Len(sMailbox) - 1)
    BoxType = RawData
    If BoxType = "Private" Then
        Open App.Path & "\Messages\" & sMailbox & "\" & gMessageID For Input As #1
    Else
        Open App.Path & "\Messages\" & gMessageID For Input As #1
    End If
    Input #1, NullString
    Input #1, NullString
    Input #1, NullString
    Input #1, FileData
    On Error Resume Next
    Do Until EOF(1)
    Input #1, tempdata
    FileData = FileData & vbCrLf & tempdata
    Loop
    Close #1
    SendClientCommand "ViewMessage" & Chr(1) & FileData, wsIndex
    Exit Sub
Case "#RemoveMessage"
    gMessageID = Left(RawData, InStr(1, RawData, Chr(1), vbTextCompare) - 1)
    RawData = Right(RawData, Len(RawData) - Len(gMessageID) - 1)
    sMailbox = Left(RawData, InStr(1, RawData, Chr(1), vbTextCompare) - 1)
    RawData = Right(RawData, Len(RawData) - Len(sMailbox) - 1)
    BoxType = RawData
    If BoxType = "Private" Then
        Kill App.Path & "\Messages\" & sMailbox & "\" & gMessageID
    Else
        Kill App.Path & "\Messages\" & gMessageID
    End If
    SendClientCommand "MessageSentOK", wsIndex
    Exit Sub
Case "#LocalFServe"
    sFilename = Left(RawData, InStr(1, RawData, Chr(1), vbTextCompare) - 1)
    sFilelen = Right(RawData, Len(RawData) - Len(sFilename) - 1)
    Set fsv = MainForm.FileList.ListItems.Add(, , sFilename)
    fsv.SubItems(1) = sFilelen
    fsv.SubItems(2) = GetUsernameFromWS(wsIndex)
    DoEvents
    SendClientCommand "SendNextFile", wsIndex
    Exit Sub
Case "#InitSearch"
    Dim SearchString As String
    SearchString = RawData
    Dim SearchClass As New UserSearch
    SearchClass.UserWS = wsIndex
    SearchClass.DoSearch SearchString
    Exit Sub
Case "#StartDownload"
    sdoFilename = Left(RawData, InStr(1, RawData, Chr(1), vbTextCompare) - 1)
    sdoHost = Right(RawData, Len(RawData) - Len(sdoFilename) - 1)
    SendClientCommand "TargetIP" & Chr(1) & MainForm.wsClient(MainForm.UserList.ListItems(LocateUserFromWS(wsIndex)).SubItems(1)).RemoteHostIP & Chr(1) & sdoFilename, wsIndex
    Exit Sub
Case "#GetRoomName"
    Dim usrRoomName As String
    usrRoomName = "(" & GetMemberCount(MainForm.Roomlist.ListItems(Val(RawData) + 1).Text) & ") " & MainForm.Roomlist.ListItems(Val(RawData) + 1).Text
    SendClientCommand "AddRoom" & Chr(1) & usrRoomName, wsIndex
    Exit Sub
Case "#JoinRoom"
    Dim sPassword As String
    Dim sRoomData As String
    sRoomData = Left(RawData, InStr(1, RawData, Chr(1), vbTextCompare) - 1)
    sPassword = Right(RawData, Len(RawData) - Len(sRoomData) - 1)
    If sPassword = "<null>" Then sPassword = ""
    UserJoin sRoomData, GetUsernameFromWS(wsIndex), wsIndex, sPassword
    Exit Sub
Case "#GetRoomMembers"
    sendingdata = AssembleMemberList(RawData)
    SendClientCommand "RecieveRoomMembers" & Chr(1) & RawData & Chr(1) & sendingdata, wsIndex
    Exit Sub
Case "#RoomPost"
    Dim sdtRoomname As String
    Dim messagedata As String
    sdtRoomname = Left(RawData, InStr(1, RawData, Chr(1), vbTextCompare) - 1)
    messagedata = Right(RawData, Len(RawData) - Len(sdtRoomname) - 1)
    PostRoomData sdtRoomname, "UpdatePost", messagedata
    Exit Sub
Case "#RoomRemove"
    '// Remove user from room
    RemoveFromRoom RawData, GetUsernameFromWS(wsIndex)
    Exit Sub
Case "#FinishedRoomJoin"
    PostRoomData RawData, "AddRoomUser", GetUsernameFromWS(wsIndex)
    Exit Sub
Case "#CreateRoom"
    Dim crRoomname As String
    Dim crRoomtype As String
    Dim crPassword As String
    Dim TempString As String
    TempString = RawData
    crRoomname = Left(TempString, InStr(1, TempString, Chr(1), vbTextCompare) - 1)
    TempString = Right(TempString, Len(TempString) - Len(crRoomname) - 1)
    crRoomtype = Left(TempString, InStr(1, TempString, Chr(1), vbTextCompare) - 1)
    TempString = Right(TempString, Len(TempString) - Len(crRoomtype) - 1)
    crPassword = TempString
    CreateRoom crRoomname, crRoomtype, wsIndex, crPassword
    Exit Sub
End Select
End Sub
