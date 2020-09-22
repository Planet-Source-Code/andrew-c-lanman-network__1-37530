Attribute VB_Name = "ShortCutAPI"
'// Shortcut Functions Module
Option Explicit
Option Base 1
Public Declare Function OSfCreateShellLink Lib "vb6stkit.dll" Alias "fCreateShellLink" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String, ByVal fPrivate As Long, ByVal sParent As String) As Long
Public Const gstrQUOTE$ = """"
Public Sub CreateShellLink(ByVal strLinkPath As String, ByVal strGroupName As String, ByVal strLinkArguments As String, ByVal strLinkName As String, ByVal fPrivate As Boolean, sParent As String, Optional ByVal fLog As Boolean = True)
Dim fSuccess As Boolean
Dim intMsgRet As Integer
Dim lREt       As Boolean
   strLinkName = strUnQuoteString(strLinkName)
   strLinkPath = strUnQuoteString(strLinkPath)
   If StrPtr(strLinkArguments) = 0 Then strLinkArguments = ""
   lREt = OSfCreateShellLink(strGroupName, strLinkName, strLinkPath, strLinkArguments, fPrivate, sParent)
End Sub

Public Function strUnQuoteString(ByVal strQuotedString As String)
    strQuotedString = Trim$(strQuotedString)
    If Mid$(strQuotedString, 1, 1) = gstrQUOTE Then
        If Right$(strQuotedString, 1) = gstrQUOTE Then
 
            strQuotedString = Mid$(strQuotedString, 2, Len(strQuotedString) - 2)
        End If
    End If
    strUnQuoteString = strQuotedString
End Function

