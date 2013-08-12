Attribute VB_Name = "读取SDR"
' SDR格式数据读取

Private Const pov As String = "B1" ' Version单元格
Private Const pos As String = "B9" ' site.site单元格
Private Const pres As String = "site" ' site表前缀
Private Const pred As String = "data" ' data表前缀

Function decSDR(rst As Object)
    Dim s As Station: Set s = New Station
    
    Dim sn As String: sn = pres & rst.Range(pos).Value
    
    If sheetExist(sn) Then
        s.setSheet Sheets(sn)
    Else
        Sheets.Add after:=Sheets(Sheets.Count)
        ActiveSheet.Name = sn
        s.newStation ActiveSheet
        
        decInfoSDR rst, s
    End If
    
    rst.Copy after:=Sheets(Sheets.Count)
    ActiveSheet.Name = pred + s.id
    Rows("1:" & (s.DataStart - 1)).Delete Shift:=xlUp
    
    adjustData ActiveSheet, s
    addStation s
End Function

Private Function decInfoSDR(rs As Object, s As Object)
    s.System = "SDR"
    s.Version = rs.Range(pov).Value

    Dim reISH As Object: Set reISH = CreateObject("vbscript.regexp")
    reISH.Pattern = "^([\d\.]+)\s*(m|ft)"

    Dim i As Single
    For i = 1 To rs.UsedRange.Rows.Count
        If InStr(1, rs.Cells(i, 1).Value, "Logger", 1) > 0 Then
            With s.Logger
                .Model = rs.Cells(i + 1, 2).Value
                .Serial = rs.Cells(i + 2, 2).Value
                .HardwareRev = rs.Cells(i + 3, 2).Value
            End With
            
            i = i + 3
        ElseIf InStr(1, Cells(i, 1).Value, "Site", 1) > 0 Then
            With s.Site
                .Site = rs.Cells(i + 1, 2).Value
                .SiteDesc = rs.Cells(i + 2, 2).Value
                .ProjectCode = rs.Cells(i + 3, 2).Value
                .ProjectDesc = rs.Cells(i + 4, 2).Value
                .SiteLocation = rs.Cells(i + 5, 2).Value
                .SiteElevation = rs.Cells(i + 6, 2).Value
                .Latitude = rs.Cells(i + 7, 2).Value
                .Longitude = rs.Cells(i + 8, 2).Value
                .TimeOffset = rs.Cells(i + 9, 2).Value
            End With

            i = i + 9
        ElseIf InStr(1, rs.Cells(i, 1).Value, "Channel", 1) > 0 Then
            Dim ss As Object: Set ss = s.newSensor
            With ss
                .channel = rs.Cells(i, 2).Value
                .cat = rs.Cells(i + 1, 2).Value
                .Description = rs.Cells(i + 2, 2).Value
                .Details = rs.Cells(i + 3, 2).Value
                .SerialNumber = rs.Cells(i + 4, 2).Value
                .ScaleFactor = rs.Cells(i + 6, 2).Value
                .Offset = rs.Cells(i + 7, 2).Value
                .Units = rs.Cells(i + 8, 2).Value
                .avg = (rs.Cells(i, 2).Value - 1) * 4 + 2
                .Sd = (rs.Cells(i, 2).Value - 1) * 4 + 3
                .Min = (rs.Cells(i, 2).Value - 1) * 4 + 4
                .Max = (rs.Cells(i, 2).Value - 1) * 4 + 5
            End With
            
            If Len(ss.channel) < 1 Then
                Error "Channel: Ch" + ss.channel
            End If
            
            Select Case ss.Units
                Case "", "-----", "unit"
                    ss.NotInstalled = True
                Case Else
                    ss.NotInstalled = False
            End Select
            
            Set mymatches = reISH.Execute(rs.Cells(i + 5, 2).Value)
            If mymatches.Count >= 1 Then
                Set mymatch = mymatches(0)
                If mymatch.SubMatches.Count >= 2 Then
                    ss.height = mymatch.SubMatches(0)
        
                    If mymatch.SubMatches(1) = "ft" Then
                        ss.height = ss.height * 0.3048
                    End If
                End If
            End If
            
            i = i + 8
        ElseIf InStr(1, rs.Cells(i, 1).Value, "Date", 1) > 0 Then
            s.DataStart = i
            Exit For
        End If
    Next i
    
    s.id = s.Site.Site
End Function

