Attribute VB_Name = "Reg"
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long

Private Const REG_SZ = 1
Private Const REG_EXPAND_SZ = 2
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const REG_MULTI_SZ = 7

Private Const KEY_QUERY_VALUE As Long = &H1
Private Const KEY_ALL_ACCESS = &H3F
Private Const REG_OPTION_NON_VOLATILE = 0

Private Const HKEY_CLASSES_ROOT As Long = &H80000000
Private Const HKEY_CURRENT_USER As Long = &H80000001
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Public Const HKCR = 0
Public Const HKLM = 1
Public Const HKCU = 2

Public Const SZ = 0
Public Const DW = 1
Public Sub CK(ByVal lRegKey As Long, ByVal lKey As String, ByVal lValueName As String, lValue As Variant, ByVal lValueType As Long)
On Error Resume Next
Dim H&
Dim Key&
Select Case lRegKey

    Case HKCR
    lRegKey = HKEY_CLASSES_ROOT
    
    Case HKLM
    lRegKey = HKEY_LOCAL_MACHINE

    Case HKCU
    lRegKey = HKEY_CURRENT_USER
    
End Select
H = RegCreateKeyEx(lRegKey, lKey, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, Key, H)
    If lValueType = SZ Then
        Dim szVal$
        szVal = lValue
        H = RegSetValueEx(Key, lValueName, 0&, REG_SZ, ByVal szVal, Len(szVal) + 1)
        Else
        Dim lVal&
        lVal = lValue
        H = RegSetValueEx(Key, lValueName, 0&, REG_DWORD, lVal, 4)
        End If
H = RegCloseKey(Key)
End Sub
Public Function RK(lRegKey As Long, ByVal lKey As String, ByVal lName As String) As Variant
On Error Resume Next
Dim Key&, zKey&, hKey$, X&
Select Case lRegKey

    Case HKCR
    lRegKey = HKEY_CLASSES_ROOT
    
    Case HKLM
    lRegKey = HKEY_LOCAL_MACHINE

    Case HKCU
    lRegKey = HKEY_CURRENT_USER
    
End Select

X = RegOpenKeyEx(lRegKey, lKey, 0, KEY_QUERY_VALUE, zKey)
X = RegQueryValueEx(zKey, lName, 0&, REG_SZ, 0&, Key)
hKey = String(Key, Chr(32))
If X <= 2 Then
    RK = ""
    Exit Function
    End If

X = RegQueryValueEx(zKey, lName, 0&, REG_SZ, ByVal hKey, Key)
hKey = Left$(hKey, Key - 1)

X = RegCloseKey(zKey)
RK = hKey
End Function
