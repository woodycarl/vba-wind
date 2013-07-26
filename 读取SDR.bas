Attribute VB_Name = "读取SDR"

' SDR 格式数据读取

Function decInfoSDR() As String
    Dim s As New Station
    Dim ss As Sensor

    s.System = "SDR"
    s.Version = Range("B1").Value
    
    Dim i As Single
    For i = 1 To ActiveSheet.UsedRange.Rows.Count
        If InStr(1, Cells(i, 1).Value, "Channel", 1) > 0 Then
            Set ss = New Sensor
            With ss
                .Channel = Cells(i, 2).Value
                .Cat = Cells(i + 1, 2).Value
                .Description = Cells(i + 2, 2).Value
                .Details = Cells(i + 3, 2).Value
                .SerialNumber = Cells(i + 4, 2).Value
                .ScaleFactor = Cells(i + 6, 2).Value
                .Offset = Cells(i + 7, 2).Value
                .Units = Cells(i + 8, 2).Value
                .Avg = (Cells(i, 2).Value - 1) * 4 + 1
                .SD = (Cells(i, 2).Value - 1) * 4 + 2
                .Min = (Cells(i, 2).Value - 1) * 4 + 3
                .Max = (Cells(i, 2).Value - 1) * 4 + 4
            End With
            
            If Len(ss.Channel) < 1 Then
                Error "传感器Channel号为空: Ch" + ss.Channel
            End If
            
            Select Case ss.Units
                Case "", "-----", "unit"
                    ss.NotInstalled = True
                Case Else
                    ss.NotInstalled = False
            End Select
            
            Set mymatches = reISH.Execute(Cells(i + 5, 2).Value)
            If mymatches.Count >= 1 Then
                Set mymatch = mymatches(0)
                If mymatch.SubMatches.Count >= 2 Then
                    ss.Height = mymatch.SubMatches(0)
        
                    If mymatch.SubMatches(1) = "ft" Then
                        ss.Height = ss.Height * 0.3048
                    End If
                End If
            End If
            
            s.SensorsR.add ss.Channel, ss

            i = i + 8
        ElseIf InStr(1, Cells(i, 1).Value, "Logger", 1) > 0 Then
            Set s.Logger = New Logger
            With s.Logger
                .Model = Cells(i + 1, 2).Value
                .Serial = Cells(i + 2, 2).Value
                .HardwareRev = Cells(i + 3, 2).Value
            End With
            
            i = i + 3
        ElseIf InStr(1, Cells(i, 1).Value, "Site", 1) > 0 Then
            Set s.Site = New Site
            With s.Site
                .Site = Cells(i + 1, 2).Value
                .SiteDesc = Cells(i + 2, 2).Value
                .ProjectCode = Cells(i + 3, 2).Value
                .ProjectDesc = Cells(i + 4, 2).Value
                .SiteLocation = Cells(i + 5, 2).Value
                .SiteElevation = Cells(i + 6, 2).Value
                .Latitude = Cells(i + 7, 2).Value
                .Longitude = Cells(i + 8, 2).Value
                .TimeOffset = Cells(i + 9, 2).Value
            End With

            i = i + 9
        ElseIf InStr(1, Cells(i, 1).Value, "Date", 1) > 0 Then
            s.DataStart = i + 1
            Exit For
        End If
    
    Next i
    
    s.id = s.Site.Site
    
    addStation s

    decInfoSDR = s.id
    
End Function

Function decDataSDR(id As String)
    ' 复制数据到新表
    
    Dim maxX, maxY
    maxX = ActiveSheet.UsedRange.Rows.Count
    maxY = ActiveSheet.UsedRange.Columns.Count
    
    Dim x As String, y As String
    x = Cells(Stations(id).DataStart - 1, 1).Address
    y = Cells(maxX, maxY).Address

    ActiveSheet.Range(x + ":" + y).Copy
    Sheets.add After:=Sheets(Sheets.Count)
    ActiveSheet.Paste
    Range("A1").Select
    
    ' 必要的调整
    
    adjustData id
End Function

