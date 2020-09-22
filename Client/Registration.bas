Attribute VB_Name = "Registration"
'// Copy Protection Functions Module
Public Function DoSerialKey()
    Dim KeyName
    For X = 1 To Len(GetComputerName)
        KeyName = KeyName & Asc(Right(Left(GetComputerName, X), 1)) * 2
    Next X
    DoSerialKey = KeyName
End Function

Public Function CheckSerialKey()
    Dim CurrentKey
    Dim MessageDialog
    MainForm.Hide
start:
    CurrentKey = GetKey("LanMan", "Settings", "Key")
    If CurrentKey <> DoSerialKey Then
        MessageDialog = MsgBox("This copy of LanMan has not been authorised!" & Chr(10) & "Access to the lanman network has been denied!", vbExclamation + vbRetryCancel, "Cannot find authorisaton key")
        If MessageDialog = vbRetry Then
            GoTo start
        Else
            End
        End If
    End If
End Function
