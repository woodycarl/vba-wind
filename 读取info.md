Sub 读取信息()
    ' Dim raw As New Collection
    
    Dim i As Integer
    For i = 1 To Sheets.Count
        If InStr(1, Sheets(i).name, "raw", 1) > 0 Then
            ' raw.Add (i)
            Call readInfo(i)
            
        End If
    Next i
End Sub

Function readInfo(index As Integer)
    Dim s As Station
    
    Sheets(index).Select
    If InStr(1, Cells(1, 1).Value, "SDR", 1) > 0 Then

        
        Call decInfoSDR
        ' MsgBox ("sdr")
    ElseIf InStr(1, Cells(1, 1).Value, "Multi-Track Export -", 1) > 0 Then
        s.System = "Nomad2"
    
        'MsgBox ("nomad")
    End If
End Function

Function decInfoSDR()
    Dim s As Station
    s.System = "SDR"
    s.Version = Cells(1, 2).Value
    
    Dim i As Integer
    For i = 1 To ActiveSheet.UsedRange.Rows.Count
        If InStr(1, Cells(i, 1).Value, "Channel", 1) > 0 Then
            Dim sensor As sensor
            With sensor
                .Channel = Cells(i, 2).Value
                .Type = Cells(i + 1, 2).Value
                .Description = Cells(i + 2, 2).Value
                .Details = Cells(i + 3, 2).Value
                .SerialNumber = Cells(i + 4, 2).Value
                .ScaleFactor = Cells(i + 6, 2).Value
                .Offset = Cells(i + 7, 2).Value
                .Units = Cells(i + 8, 2).Value
            End With
            
            If Len(sensor.Channel) < 1 Then
                Call Error("传感器Channel号为空")
            End If
            
        End If
    
    
    Next i
    
End Function
