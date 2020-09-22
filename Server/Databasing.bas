Attribute VB_Name = "FServe"
'// Database access module
Public Const fKB = 1048576
Public Const fMB = 1073741824
Public Const fGB = 1099511627776#
Public Const fTB = 1.12589990684262E+15

Public Sub RefreshTotalNetHost()
Dim TotalSize As Long
TotalSize = 0
For x = 1 To MainForm.FileList.ListItems.Count
TotalSize = TotalSize + (MainForm.FileList.ListItems(x).SubItems(1))
Next x
MainForm.lblTotalFiles.Caption = "Total Network Capacity: " & FileSizeConv(Str(TotalSize))
End Sub
