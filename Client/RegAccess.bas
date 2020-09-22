Attribute VB_Name = "Module1"
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type
Const REG_SZ = 1
Const REG_EXPAND_SZ = 2
Const REG_DWORD = 4
Const REG_OPTION_NON_VOLATILE = 0
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Const KEY_EXECUTE = KEY_READ
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
Const ERROR_NONE = 0
Const ERROR_BADKEY = 2
Const ERROR_ACCESS_DENIED = 8
Const ERROR_SUCCESS = 0
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004

Public Function UpdateKey(KeyRoot As Long, Keyname As String, SubKeyName As String, SubKeyValue As String) As Boolean
    Dim rc As Long
    Dim hKey As Long
    Dim hDepth As Long
    Dim lpAttr As SECURITY_ATTRIBUTES
    lpAttr.nLength = 50
    lpAttr.lpSecurityDescriptor = 0
    lpAttr.bInheritHandle = True
    rc = RegCreateKeyEx(KeyRoot, Keyname, _
                        0, REG_SZ, _
                        REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpAttr, _
                        hKey, hDepth)
    
    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError
    If (SubKeyValue = "") Then SubKeyValue = " "
    rc = RegSetValueEx(hKey, SubKeyName, _
                       0, REG_SZ, _
                       SubKeyValue, LenB(StrConv(SubKeyValue, vbFromUnicode)))
                       
    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError
    rc = RegCloseKey(hKey)
    UpdateKey = True
    Exit Function
CreateKeyError:
    UpdateKey = False
    rc = RegCloseKey(hKey)
End Function

Public Function GetKeyValue(KeyRoot As Long, Keyname As String, SubKeyRef As String) As String
    Dim i As Long
    Dim rc As Long
    Dim hKey As Long
    Dim hDepth As Long
    Dim sKeyVal As String
    Dim lKeyValType As Long
    Dim tmpVal As String
    Dim KeyValSize As Long
    rc = RegOpenKeyEx(KeyRoot, Keyname, 0, KEY_ALL_ACCESS, hKey)
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
    tmpVal = String$(1024, 0)
    KeyValSize = 1024
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         lKeyValType, tmpVal, KeyValSize)
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
    tmpVal = Left$(tmpVal, InStr(tmpVal, Chr(0)) - 1)
    Select Case lKeyValType
    Case REG_SZ, REG_EXPAND_SZ
        sKeyVal = tmpVal
    Case REG_DWORD
        For i = Len(tmpVal) To 1 Step -1
            sKeyVal = sKeyVal + Hex(Asc(Mid(tmpVal, i, 1)))
        Next
        sKeyVal = Format$("&h" + sKeyVal)
    End Select
    
    GetKeyValue = sKeyVal
    rc = RegCloseKey(hKey)
    Exit Function
    
GetKeyError:
    GetKeyValue = vbNullString
    rc = RegCloseKey(hKey)
End Function

Public Sub SaveKey(AppName As String, RootName As String, Keyname As String, KeyValue As String)
UpdateKey HKEY_CURRENT_USER, "Software\Lanman\" & RootName, Keyname, KeyValue
End Sub

Public Function GetKey(AppName As String, RootName As String, Keyname As String, Optional sDefault As String)
GetKey = GetKeyValue(HKEY_CURRENT_USER, "Software\Lanman\" & RootName, Keyname)
If GetKey = vbNullString Then
   GetKey = sDefault
End If
End Function
