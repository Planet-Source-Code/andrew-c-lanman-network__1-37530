Attribute VB_Name = "MessageBoard"
Dim sTo As String
Dim sFrom As String
Dim sSubject As String
Dim sBody As String
Dim Priority As String
Public Sub UpdateCounts()
MainForm.TempFile.Path = App.Path & "\Messages\Public"
MainForm.TempFile.Refresh
MainForm.lblPublicMSGCount.Caption = "Public Messages: " & MainForm.TempFile.ListCount

Dim Counter As Integer
Counter = 0
MainForm.TempDir.Path = App.Path & "\Messages"
For x = 0 To MainForm.TempDir.ListCount - 1
If MainForm.TempDir.List(x) = App.Path & "\Messages\Public" Then GoTo SkipFolder
MainForm.TempFile.Path = MainForm.TempDir.List(x)
Counter = Counter + MainForm.TempFile.ListCount
SkipFolder:
Next x
MainForm.lblPrivateMSGCount.Caption = "Private Messages: " & Counter
End Sub


Public Sub UpdateInboxs()
MainForm.pInboxDir.Path = App.Path & "\Messages"
For x = 0 To MainForm.pInboxDir.ListCount - 1
MainForm.pInboxList.AddItem Right(MainForm.pInboxDir.List(x), Len(MainForm.pInboxDir.List(x)) - InStrRev(MainForm.pInboxDir.List(x), "\", Len(MainForm.pInboxDir.List(x)), vbTextCompare))
Next x
End Sub

Public Sub ProcessPrivateMessage(RawData As String, wsIndex As Integer)
Dim StringData As String
StringData = RawData

Priority = Left(StringData, InStr(1, StringData, Chr(1), vbTextCompare) - 1)
StringData = Right(StringData, Len(StringData) - Len(Priority) - 1)

sFrom = Left(StringData, InStr(1, StringData, Chr(1), vbTextCompare) - 1)
StringData = Right(StringData, Len(StringData) - Len(sFrom) - 1)

sTo = Left(StringData, InStr(1, StringData, Chr(1), vbTextCompare) - 1)
StringData = Right(StringData, Len(StringData) - Len(sTo) - 1)

sSubject = Left(StringData, InStr(1, StringData, Chr(1), vbTextCompare) - 1)
StringData = Right(StringData, Len(StringData) - Len(sSubject) - 1)

sBody = StringData

MainForm.TempFile.Path = App.Path & "\Accounts"
MainForm.TempFile.Refresh
For x = 0 To MainForm.TempFile.ListCount - 1
If LCase(sTo) = LCase(Left(MainForm.TempFile.List(x), Len(MainForm.TempFile.List(x)) - 4)) Then
   GoTo AccountFound
End If
Next x
SendClientCommand "MSGBadUser", wsIndex
Exit Sub

AccountFound:
Open App.Path & "\Messages\" & sTo & "\" & CreateFilename For Output As #1
Print #1, "Post Date: " & Date$
Print #1, "Post Time: " & Format(Now, "hh:mm:ss AM/PM")
Print #1, "Tag: " & Priority
Print #1, "To: " & sTo
Print #1, "From:" & sFrom
Print #1, "Subject: " & sSubject
Print #1, ""
Print #1, sBody
Close #1

SendClientCommand "MessageSentOK", wsIndex
End Sub

Public Function CreateFilename() As String
Dim sFilename
TryAgain:
sFilename = ""
Randomize
sFilename = Hex((Int(Rnd * 250)))
For x = 0 To 7
Randomize
sFilename = sFilename & Hex((Int(Rnd * 250)))
Next x
MainForm.TempDir.Path = App.Path & "\Messages\" & sTo
MainForm.Refresh
For x = 0 To MainForm.TempDir.ListCount - 1
If LCase(sFilename) = LCase(MainForm.TempDir.List(x)) Then GoTo TryAgain
Next x
Found:
CreateFilename = sFilename
End Function
