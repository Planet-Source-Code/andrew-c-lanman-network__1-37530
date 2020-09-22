VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// User search class
Public UserWS As Integer
Public SentData As Boolean
Public Sub DoSearch(SearchString As String)
On Error GoTo ExitSearch
Dim CurrentSearchIndex As Integer
MainForm.UserList.ListItems(LocateUserFromWS(UserWS)).SubItems(2) = "False"

' Perform a check. If it matches then send the data
For x = 1 To MainForm.FileList.ListItems.Count
If MainForm.UserList.ListItems(LocateUserFromWS(UserWS)).SubItems(2) = "Kill" Then
   MainForm.UserList.ListItems(LocateUserFromWS(UserWS)).SubItems(2) = "OK"
   GoTo ExitSearch
End If
If InStr(1, MainForm.FileList.ListItems(x).Text, SearchString, vbTextCompare) <> 0 Then
   SendClientCommand "SearchData" & Chr(1) & MainForm.FileList.ListItems(x) & Chr(1) & MainForm.FileList.ListItems(x).SubItems(1) & Chr(1) & MainForm.FileList.ListItems(x).SubItems(2), UserWS
   DoEvents
   SentData = True
   GoTo SkipAdv1
End If

' Perform a more detailed check, then send details
If InStr(1, SearchString, " ", vbTextCompare) = 0 Then GoTo SkipAdv1
Dim DataBank As String
Dim Individ As String
DataBank = SearchString
Do Until DataBank = "" Or DataBank = " "
Individ = Left(DataBank, InStr(1, DataBank, " ", vbTextCompare) - 1)
If InStr(1, MainForm.FileList.ListItems(x).Text, Individ, vbTextCompare) <> 0 Then
   SendClientCommand "SearchData" & Chr(1) & MainForm.FileList.ListItems(x) & Chr(1) & MainForm.FileList.ListItems(x).SubItems(1) & Chr(1) & MainForm.FileList.ListItems(x).SubItems(2), UserWS
   DoEvents
   SentData = True
End If
DataBank = Right(DataBank, Len(DataBank) - Len(Individ) - 1)
Loop
SkipAdv1:

If SentData = False Then GoTo SkipEvents
Do Until MainForm.UserList.ListItems(LocateUserFromWS(UserWS)).SubItems(2) = "OK"
DoEvents
If MainForm.UserList.ListItems(LocateUserFromWS(UserWS)).SubItems(2) = "Kill" Then
   MainForm.UserList.ListItems(LocateUserFromWS(UserWS)).SubItems(2) = "OK"
   GoTo ExitSearch
End If
Loop

SkipEvents:
Next x
ExitSearch:
MainForm.UserList.ListItems(LocateUserFromWS(UserWS)).SubItems(2) = "OK"
SendClientCommand "SearchComplete", UserWS
End Sub
