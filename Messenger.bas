Attribute VB_Name = "Messenger"
Private C   As String
Private Z   As String
Private X()
Public Function Ec(ByVal lText As String) As String
On Error Resume Next

For i = 1 To Len(lText)
    C = Mid(lText, i, 1)
    
    Z = Asc(C)
    
    Ec = Ec + Chr$(Z Xor 216)
Next

End Function
