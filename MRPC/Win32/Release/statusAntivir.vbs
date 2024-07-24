Dim av_searcher As New ManagementObjectSearcher("root\SecurityCenter2", "SELECT * FROM AntivirusProduct")
        For Each info As ManagementObject In av_searcher.Get()
            tbAv.AppendText(info.Properties("displayName").Value.ToString() & vbCrLf)

            Dim AvStatus = Hex(info.Properties("ProductState").Value.ToString())
            If Mid(AvStatus, 2, 2) = "10" Or Mid(AvStatus, 2, 2) = "11" Then
                tbAvStatus.AppendText("AntiVirus enabled" & vbCrLf)
            ElseIf Mid(AvStatus, 2, 2) = "00" Or Mid(AvStatus, 2, 2) = "01" Then
                tbAvStatus.AppendText("AntiVirus disabled" & vbCrLf)
            End If

            Dim AvCurrent = Hex(info.Properties("ProductState").Value.ToString())
            If Mid(AvStatus, 4, 2) = "00" Then
                tbAvCurrent.AppendText("AntiVirus up-to-date" & vbCrLf)
            ElseIf Mid(AvStatus, 4, 2) = "10" Then
                tbAvCurrent.AppendText("AntiVirus outdated" & vbCrLf)
            End If

        Next info