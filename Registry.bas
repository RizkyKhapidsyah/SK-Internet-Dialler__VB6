Attribute VB_Name = "Registry"

' -----------------
' ADVAPI32
' -----------------
' function prototypes, constants, and type definitions
' for Windows 32-bit Registry API

Option Explicit
'Constants
Public Const REG_SZ = 1
Public Const REG_DWORD = 4
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_DYN_DATA = &H80000006
Public Const ERROR_NONE = 0
Public Const ERROR_SUCCESS = 0
Public Const KEY_ALL_ACCESS = &H3F
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const REG_OPTION_NON_VOLATILE = 0

'Registry functions

Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long

Public Sub CreateKey(ByVal lRootKey As Long, ByVal sKeyName As String)
On Error Resume Next
'create a new registry key
Dim hKey As Long
    Call RegCreateKey(lRootKey, sKeyName, hKey)
    RegCloseKey (hKey)
End Sub
Public Sub WriteString(ByVal lRootKey As Long, ByVal sPath As String, ByVal sValueName As String, sValueData As String)
On Error Resume Next
'write a string of data to the registry
Dim hKey As Long, lResult As Long
    lResult = RegOpenKeyEx(lRootKey, sPath, vbNull, KEY_SET_VALUE, hKey)
    If lResult = ERROR_SUCCESS Then
        sValueData = sValueData & Chr(0)
        Call RegSetValueEx(hKey, sValueName, vbNull, REG_SZ, ByVal sValueData, Len(sValueData))
        Call RegCloseKey(hKey)
    End If
End Sub

Public Function ReadString(ByVal lRootKey As Long, ByVal sPath As String, ByVal sValueName As String, ByRef sDefault As String) As String
On Error Resume Next
'read a string of data from the registry

Dim hKey As Long, lResult As Long, lValueType As Long, lDataBufSize As Long, strBuf As String
    lResult = RegOpenKeyEx(lRootKey, sPath, 0, KEY_QUERY_VALUE, hKey)
    If lResult = ERROR_SUCCESS Then
        lResult = RegQueryValueEx(hKey, sValueName, vbNull, lValueType, ByVal 0&, lDataBufSize)
        If lValueType = REG_SZ Then
            strBuf = Space$(lDataBufSize)
            lResult = RegQueryValueEx(hKey, sValueName, vbNull, REG_SZ, ByVal strBuf, lDataBufSize)
            If lResult = ERROR_SUCCESS Then
                ReadString = ClearNulls(strBuf)
            Else
                ReadString = sDefault
            End If
        Else
            ReadString = sDefault
        End If
    End If
    Call RegCloseKey(hKey)
End Function
Public Sub WriteInteger(ByVal lRootKey As Long, ByVal sPath As String, ByVal sValueName As String, ByVal iValue As Integer)
On Error Resume Next
'write an int to the registry

Dim hKey As Long, lResult As Long
    lResult = RegOpenKeyEx(lRootKey, sPath, vbNull, KEY_SET_VALUE, hKey)
    If lResult = ERROR_SUCCESS Then
        Call RegSetValueEx(hKey, sValueName, vbNull, REG_DWORD, CLng(iValue), 4)
        Call RegCloseKey(hKey)
    End If
End Sub
Public Function ReadInteger(ByVal lRootKey As Long, ByVal sPath As String, ByVal sValueName As String, iDefault As Integer) As Integer
On Error Resume Next
'read an int from the registry

Dim hKey As Long, lResult As Long, lValueType As Long, lData As Long
    lResult = RegOpenKeyEx(lRootKey, sPath, 0, KEY_QUERY_VALUE, hKey)
    If lResult = ERROR_SUCCESS Then
        lResult = RegQueryValueEx(hKey, sValueName, 0&, lValueType, lData, LenB(lData))
        If lResult = ERROR_SUCCESS And lValueType = REG_DWORD Then
            ReadInteger = CInt(lData)
        Else
            ReadInteger = iDefault
        End If
        Call RegCloseKey(hKey)
    Else
        ReadInteger = iDefault
    End If
End Function
Public Function KeyExists(ByVal lRootKey As Long, ByVal strKeyName As String) As Boolean
On Error Resume Next
'check if the given key exists in the registry

Dim hKey As Long, lResult As Long
    lResult = RegOpenKeyEx(lRootKey, strKeyName, 0, 0&, hKey)
    Call RegCloseKey(hKey)
    KeyExists = (lResult = ERROR_SUCCESS)
End Function
Public Function ReadLong(ByVal lRootKey As Long, strPath As String, strValueName As String, lDefault As Long) As Long
On Error Resume Next
'read a long from the registry

Dim hKey As Long, lResult As Long, lData As Long
    lResult = RegOpenKeyEx(lRootKey, strPath, 0, KEY_QUERY_VALUE, hKey)
    If lResult = ERROR_SUCCESS Then
        lResult = RegQueryValueEx(hKey, strValueName, 0&, REG_DWORD, lData, LenB(lData))
        If lResult = ERROR_SUCCESS Then
            ReadLong = lData
        Else
            ReadLong = lDefault
        End If
        Call RegCloseKey(hKey)
    End If
End Function
Public Function ValueExists(ByVal lRootKey As Long, ByVal lpSubKey As String, ByVal strValueName As String) As Boolean
On Error Resume Next
'check if a value exists the registry
Dim hKey As Long, lResult As Long
lResult = RegOpenKeyEx(lRootKey, lpSubKey, 0, KEY_QUERY_VALUE, hKey)
    If lResult = ERROR_SUCCESS Then
        lResult = RegQueryValueEx(hKey, strValueName, 0&, 0&, ByVal 0&, ByVal 0&)
        Call RegCloseKey(hKey)
    End If
    If lResult = ERROR_SUCCESS Then ValueExists = True
End Function
Public Sub DeleteValue(ByVal lRootKey As Long, ByVal strPath As String, ByVal sValueName As String)
Dim hKey As Long, lResult As Long
    lResult = RegOpenKeyEx(lRootKey, strPath, 0&, KEY_SET_VALUE, hKey)
    If lResult = ERROR_SUCCESS Then
        Call RegDeleteValue(hKey, sValueName)
        Call RegCloseKey(hKey)
    End If
End Sub
Sub CreateNewKey(sNewKeyName As String, lPredefinedKey As Long)
Dim hNewKey As Long
Dim lRetVal As Long
lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, _
vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
0&, hNewKey, lRetVal)
RegCloseKey (hNewKey)
End Sub
Public Sub SetKeyValue(sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As Long)
Dim Zero As Long, IRetVal As Long, hKey As Long
IRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKeyName, Zero, KEY_ALL_ACCESS, hKey)
If IRetVal Then MsgBox "RegOpenKey error - " & IRetVal
IRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
If IRetVal Then MsgBox "SetValue error - " & IRetVal
RegCloseKey (hKey)
End Sub
Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
Dim lValue As Long
Dim sValue As String
Select Case lType
Case REG_SZ
sValue = vValue & Chr$(0)
SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
Case REG_DWORD
lValue = vValue
SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
End Select
End Function
