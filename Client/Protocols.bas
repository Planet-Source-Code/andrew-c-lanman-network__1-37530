Attribute VB_Name = "ProtocolClass"
Public MessageCount As Integer
Public CounterBuffer As Integer
Public IndexHeaders(0 To 1000)
Public Sub SendServerCommand(StringData As String)
On Error Resume Next
MainForm.Server.SendData "#" & StringData
End Sub

Public Sub ProcessCommand(StringData As String)
On Error Resume Next
Select Case StringData
'
Case "#SendLogin"
    SendServerCommand "Login" & Chr(1) & frmLogin.Username.Text & Chr(1) & frmLogin.Password.Text
    Exit Sub
Case "#AccountExists"
    MsgBox "User account " & frmCreateLogin.Username.Text & " already exists on server!", vbInformation, "Account creation failed!"
    Exit Sub
Case "#CreatedOK"
    MsgBox "User account " & frmCreateLogin.Username.Text & " successfully created!", vbInformation, "Account created!"
    frmCreateLogin.Hide
    frmLogin.Username.Text = frmCreateLogin.Username.Text
    frmLogin.Password.Text = frmCreateLogin.Password.Text
    SendServerCommand "Login" & Chr(1) & frmLogin.Username.Text & Chr(1) & frmLogin.Password.Text
    Exit Sub
Case "#NoAccount"
    MsgBox "User account " & frmLogin.Username.Text & " does not exist on server!", vbInformation, "Login failed!"
    frmLogin.Username.Text = ""
    frmLogin.Password.Text = ""
    Exit Sub
Case "#LoginOK"
    If AppSettings.Startmin.Value = 1 Then
        MainForm.Hide
    Else
        MainForm.Show
    End If
    frmLogin.Hide
    Startup.Hide
    frmCreateLogin.Hide
    LogOn
    SaveKey "Lanman", "Settings", "Username", frmLogin.Username.Text
    If frmLogin.Check1.Value = 1 Then
       AppSettings.AutoLogon.Value = 1
       SaveKey "Lanman", "Settings", "AutoLogon", "1"
       SaveKey "Lanman", "Settings", "Password", frmLogin.Password.Text
    End If
    cLocalSender = 0
    If frmFServe.LocalFiles.ListCount = 0 Then Exit Sub
    SendServerCommand "LocalFServe" & Chr(1) & frmFServe.LocalFiles.List(cLocalSender) & Chr(1) & FileLen("C:\Shared Folder\" & frmFServe.LocalFiles.List(cLocalSender))
    Exit Sub
Case "#BadPassword"
    MsgBox "The password for " & frmLogin.Username.Text & " is incorrect!", vbInformation, "Login failed!"
    frmLogin.Show
    frmLogin.Password.SetFocus
    frmLogin.Password.SelStart = 0
    frmLogin.Password.SelLength = Len(frmLogin.Password.Text)
    Exit Sub
Case "#MSGBadUser"
    MsgBox "The user " & frmCompose.txtTo.Text & " does not have an account on this server!", vbInformation, "Post Failed!"
    frmCompose.txtTo.SelStart = 0
    frmCompose.txtTo.SelLength = Len(frmCompose.txtTo.Text)
    Exit Sub
Case "#MessageSentOK"
    frmCompose.Hide
    frmCompose.MessageBody.Text = ""
    frmCompose.Subject.Text = ""
    frmCompose.txtTo.Text = ""
    MessageBoard.SetFolder.Caption = "View Private"
    MessageBoard.Folder = "Private"
    Call MessageBoard.SetFolder_Click
    Exit Sub
Case "#SendNextFile"
    cLocalSender = cLocalSender + 1
    If cLocalSender >= frmFServe.LocalFiles.ListCount Then
        SendServerCommand "DoneFileList"
        Exit Sub
    End If
    SendServerCommand "LocalFServe" & Chr(1) & frmFServe.LocalFiles.List(cLocalSender) & Chr(1) & FileLen("C:\Shared Folder\" & frmFServe.LocalFiles.List(cLocalSender))
    Exit Sub
Case "#SearchComplete"
    frmFServe.Caption = "Lanman File Server"
    frmFServe.StatusBar1.SimpleText = "Search returned " & frmFServe.SearchList.ListItems.Count & " results."
    frmFServe.Searching = False
    frmFServe.cmdGo.Caption = "&Go"
    Exit Sub
Case "#AlreadyOnline"
    MainForm.ConnectRetry.Enabled = False
    MsgBox "Someone is already logged-in by this account.", vbExclamation, "Login failed"
    LogOff
    Exit Sub
End Select
Select Case Left(StringData, InStr(1, StringData, Chr(1), vbTextCompare) - 1)
'
Case "#PrivateHeaders"
    On Error Resume Next
    Dim HeaderData As String
    Dim HeaderCount As String
    Dim rawdata As String
    rawdata = Right(StringData, Len(StringData) - 16)
    HeaderCount = Left(rawdata, InStr(1, rawdata, Chr(1), vbTextCompare) - 1)
    HeaderData = Right(rawdata, Len(rawdata) - Len(HeaderCount) - 1)
    rawdata = HeaderData
    For X = 0 To HeaderCount - 1
    If InStr(1, rawdata, Chr(1), vbTextCompare) = 0 Then
       MessageBoard.LoadIntoPrivate rawdata, Val(X)
    Else
        MessageBoard.LoadIntoPrivate Left(rawdata, InStr(1, rawdata, Chr(1), vbTextCompare) - 1), Val(X)
        rawdata = Right(rawdata, Len(rawdata) - InStr(1, rawdata, Chr(1), vbTextCompare))
    End If
    Next X
    MessageCount = HeaderCount
    CounterBuffer = 0
    SendServerCommand "GetMessageTitle" & Chr(1) & MessageBoard.GetHeaderDetail(Val(CounterBuffer)) & Chr(1) & frmLogin.Username.Text
    Exit Sub
Case "#PublicHeaders"
    On Error Resume Next
    rawdata = Right(StringData, Len(StringData) - 15)
    HeaderCount = Left(rawdata, InStr(1, rawdata, Chr(1), vbTextCompare) - 1)
    HeaderData = Right(rawdata, Len(rawdata) - Len(HeaderCount) - 1)
    rawdata = HeaderData
    For X = 0 To HeaderCount - 1
    If InStr(1, rawdata, Chr(1), vbTextCompare) = 0 Then
       MessageBoard.LoadIntoPrivate rawdata, Val(X)
    Else
        MessageBoard.LoadIntoPrivate Left(rawdata, InStr(1, rawdata, Chr(1), vbTextCompare) - 1), Val(X)
        rawdata = Right(rawdata, Len(rawdata) - InStr(1, rawdata, Chr(1), vbTextCompare))
    End If
    Next X
    MessageCount = HeaderCount
    CounterBuffer = 0
    SendServerCommand "GetPublicTitle" & Chr(1) & MessageBoard.GetHeaderDetail(Val(CounterBuffer))
    Exit Sub
Case "#AddMessage"
    Dim SentDate As String
    Dim SentTime As String
    Dim SentFrom As String
    Dim sPriority
    Dim Subject
    Dim cMessageID As String
    rawdata = Right(StringData, Len(StringData) - 12)
    cMessageID = Left(rawdata, InStr(1, rawdata, Chr(1), vbTextCompare) - 1)
    rawdata = Right(rawdata, Len(rawdata) - Len(cMessageID) - 1)
    sPriority = Left(rawdata, InStr(1, rawdata, Chr(1), vbTextCompare) - 1)
    rawdata = Right(rawdata, Len(rawdata) - Len(sPriority) - 1)
    SentDate = Left(rawdata, InStr(1, rawdata, Chr(1), vbTextCompare) - 1)
    rawdata = Right(rawdata, Len(rawdata) - Len(SentDate) - 1)
    SentTime = Left(rawdata, InStr(1, rawdata, Chr(1), vbTextCompare) - 1)
    rawdata = Right(rawdata, Len(rawdata) - Len(SentTime) - 1)
    SentFrom = Left(rawdata, InStr(1, rawdata, Chr(1), vbTextCompare) - 1)
    rawdata = Right(rawdata, Len(rawdata) - Len(SentFrom) - 1)
    Subject = rawdata
    Set h = MessageBoard.Messages.ListItems.Add(, cMessageID, SentDate & " / " & SentTime & "   -" & SentFrom & ": " & Subject, Val(sPriority), Val(sPriority))
    CounterBuffer = CounterBuffer + 1
    If CounterBuffer >= MessageCount Then Exit Sub
    SendServerCommand "GetPublicTitle" & Chr(1) & MessageBoard.GetHeaderDetail(CounterBuffer) & Chr(1) & frmLogin.Username.Text
    Exit Sub
Case "#ViewMessage"
    Dim FileData As String
    FileData = Right(StringData, Len(StringData) - 13)
    frmMSGView.MessageData.Text = FileData
    frmMSGView.Show
    Exit Sub
Case "#SearchData"
    Dim sdFilename, sdFilesize, sdHost As String
    rawdata = Right(StringData, Len(StringData) - 12)
    sdFilename = Left(rawdata, InStr(1, rawdata, Chr(1), vbTextCompare) - 1)
    rawdata = Right(rawdata, Len(rawdata) - Len(sdFilename) - 1)
    sdFilesize = Left(rawdata, InStr(1, rawdata, Chr(1), vbTextCompare) - 1)
    rawdata = Right(rawdata, Len(rawdata) - Len(sdFilesize) - 1)
    sdHost = rawdata
    Set sh = frmFServe.SearchList.ListItems.Add(, , sdFilename)
    sh.SubItems(1) = sdFilesize
    sh.SubItems(2) = sdHost
    DoEvents
    SendServerCommand "SearchProceed"
    DoEvents
    Exit Sub
Case "#SetRoomCount"
    rawdata = Right(StringData, Len(StringData) - 14)
    RoomList.NumberofRooms = rawdata
    RoomList.RoomCounter = 0
    DoEvents
    SendServerCommand "GetRoomName" & Chr(1) & RoomList.RoomCounter
    Exit Sub
Case "#AddRoom"
    rawdata = Right(StringData, Len(StringData) - 9)
    RoomList.RoomList.AddItem rawdata
    RoomList.RoomCounter = RoomList.RoomCounter + 1
    If RoomList.RoomCounter = RoomList.NumberofRooms Then
       Exit Sub
    End If
    SendServerCommand "GetRoomName" & Chr(1) & RoomList.RoomCounter
    Exit Sub
Case "#BadRoom"
    rawdata = Right(StringData, Len(StringData) - 9)
    MsgBox "Server was unable to join you to the '" & rawdata & "' chatroom!" & vbCrLf & "The chatroom name is invalid!", vbInformation, "Failed to join"
    Exit Sub
Case "#RoomJoin"
    rawdata = Right(StringData, Len(StringData) - 10)
    If FindFreeHandle = -1 Then
        MsgBox "All available chat windows are being used!" & vbCrLf & "Close some open chatrooms!", vbExclamation, "Cannot open more than 20 rooms"
        Exit Sub
    End If
    Set ChatRooms.WindowBuffers(FindFreeHandle) = New ChatWnd
    ChatRooms.WindowBuffers(FindFreeHandle).Show
    SkinWindow ChatRooms.WindowBuffers(FindFreeHandle)
    RoomPass.Hide
    CreateRoom.Hide
    ChatRooms.WindowBuffers(FindFreeHandle).TitleText = "Now talking in " & rawdata
    ChatRooms.WindowBuffers(FindFreeHandle).Tag = rawdata
    ChatRooms.RoomMatches(FindFreeHandle) = rawdata
    ChatRooms.Checkers(FindFreeHandle) = True
    SendServerCommand "GetRoomMembers" & Chr(1) & rawdata
    Exit Sub
Case "#RecieveRoomMembers"
    Dim sjRoomname As String
    rawdata = Right(StringData, Len(StringData) - 20)
    sjRoomname = Left(rawdata, InStr(1, rawdata, Chr(1), vbTextCompare) - 1)
    memberdata = Right(rawdata, Len(rawdata) - Len(sjRoomname) - 1)
    
    Do Until memberdata = ""
    If InStr(1, memberdata, Chr(1), vbTextCompare) <> 0 Then
        tempdata = Left(memberdata, InStr(1, memberdata, Chr(1), vbTextCompare) - 1)
        memberdata = Right(memberdata, Len(memberdata) - Len(tempdata) - 1)
        ChatRooms.WindowBuffers(ChatRooms.GetWndhandle(sjRoomname)).UserList.AddItem tempdata
    Else
        tempdata = memberdata
        ChatRooms.WindowBuffers(ChatRooms.GetWndhandle(sjRoomname)).UserList.AddItem tempdata
        GoTo ExitLoop
    End If
    Loop
ExitLoop:
    SendServerCommand "FinishedRoomJoin" & Chr(1) & sjRoomname
    RoomList.Hide
    Exit Sub
Case "#UpdatePost"
    Dim sjtRoomname As String
    Dim sjtPostData As String
    rawdata = Right(StringData, Len(StringData) - 12)
    sjtRoomname = Left(rawdata, InStr(1, rawdata, Chr(1), vbTextCompare) - 1)
    sjtPostData = Right(rawdata, Len(rawdata) - Len(sjtRoomname) - 1)
    If WindowBuffers(GetWndhandle(sjtRoomname)).Conversation.Text = "" Then
       WindowBuffers(GetWndhandle(sjtRoomname)).Conversation.Text = sjtPostData
       Exit Sub
    Else
       WindowBuffers(GetWndhandle(sjtRoomname)).Conversation.Text = WindowBuffers(GetWndhandle(sjtRoomname)).Conversation.Text & vbCrLf & sjtPostData
       Exit Sub
    End If
Case "#AddRoomUser"
    On Error Resume Next
    Dim aruRoomname As String
    Dim aruPostData As String
    rawdata = Right(StringData, Len(StringData) - 13)
    aruRoomname = Left(rawdata, InStr(1, rawdata, Chr(1), vbTextCompare) - 1)
    aruPostData = Right(rawdata, Len(rawdata) - Len(aruRoomname) - 1)
    For X = 0 To WindowBuffers(GetWndhandle(aruRoomname)).UserList.ListCount - 1
    If WindowBuffers(GetWndhandle(aruRoomname)).UserList.List(X) = aruPostData Then
        GoTo Skipped
    End If
    Next X
    WindowBuffers(GetWndhandle(aruRoomname)).UserList.AddItem aruPostData
Skipped:
    If WindowBuffers(GetWndhandle(aruRoomname)).Conversation.Text = "" Then
       WindowBuffers(GetWndhandle(aruRoomname)).Conversation.Text = "  *User " & aruPostData & " has joined " & WindowBuffers(GetWndhandle(aruRoomname)).Tag
    Else
       WindowBuffers(GetWndhandle(aruRoomname)).Conversation.Text = WindowBuffers(GetWndhandle(aruRoomname)).Conversation.Text & vbCrLf & "  *User " & aruPostData & " has joined " & WindowBuffers(GetWndhandle(aruRoomname)).Tag
    End If
    Exit Sub
Case "#UserRoomRemove"
    On Error Resume Next
    Dim usmRawdata As String
    Dim usmRoomname As String
    Dim usmPostData As String
    usmRawdata = Right(StringData, Len(StringData) - 16)
    usmRoomname = Left(usmRawdata, InStr(1, usmRawdata, Chr(1), vbTextCompare) - 1)
    usmPostData = Right(usmRawdata, Len(usmRawdata) - Len(usmRoomname) - 1)
    For X = 0 To WindowBuffers(GetWndhandle(usmRoomname)).UserList.ListCount - 1
    If WindowBuffers(GetWndhandle(usmRoomname)).UserList.List(X) = usmPostData Then
        WindowBuffers(GetWndhandle(usmRoomname)).UserList.RemoveItem X
        If WindowBuffers(GetWndhandle(usmRoomname)).Conversation.Text = "" Then
            WindowBuffers(GetWndhandle(usmRoomname)).Conversation.Text = "  *User " & aruPostData & " has quit " & WindowBuffers(GetWndhandle(usmRoomname)).Tag
        Else
            WindowBuffers(GetWndhandle(usmRoomname)).Conversation.Text = WindowBuffers(GetWndhandle(usmRoomname)).Conversation.Text & vbCrLf & "  *User " & usmPostData & " has quit " & WindowBuffers(GetWndhandle(usmRoomname)).Tag
        End If
        Exit Sub
    End If
    Next X
    Exit Sub
Case "#BadRoomPass"
    rawdata = Right(StringData, Len(StringData) - 13)
    If RoomPass.Counter <> 1 Then
       RoomPass.PasswordText.Caption = "Enter password for the " & rawdata & " private chat-room:"
       RoomPass.Tag = rawdata
       RoomPass.Show
       RoomPass.AccessDeniedTMR.Enabled = False
       RoomPass.txtPassword.SetFocus
    Else
       RoomPass.AccessDeniedTMR.Enabled = True
       Call RoomPass.AccessDeniedTMR_Timer
    End If
    Exit Sub
End Select
End Sub
