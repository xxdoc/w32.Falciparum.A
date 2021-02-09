Attribute VB_Name = "FileXOperation"
Option Explicit

Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Public Sub CF(ByVal ORX$, ByVal DSX$)
On Error Resume Next

CopyFile ORX, DSX, 0

End Sub
Public Sub MF(ByVal ORX$, ByVal DSX$)
On Error Resume Next

MoveFile ORX, DSX

End Sub
Public Sub DF(ByVal DSX$)
On Error Resume Next

DeleteFile DSX

End Sub
