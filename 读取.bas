Attribute VB_Name = "读取"
Public Type NomadSensor
    Height      As Single
    Description As String
    Units       As String
    Cat         As String
End Type

Private reISH As Object ' info sensor height
Private re1 As Object   ' 判断时间正则1:
Private re2 As Object   ' 判断时间正则2:



Sub 读取数据()
    initRead

    Dim i As Integer
    For i = 1 To Sheets.Count
        If InStr(1, Sheets(i).Name, "raw", 1) > 0 Then
            Sheets(i).Select

            decRaw
            
        End If
    Next i

End Sub

' 设定正则
Function initRead()
    Set reISH = CreateObject("vbscript.regexp")
    With reISH
       .Global = True
       .Pattern = "^([\d\.]+)\s*(m|ft)"
    End With

    Set re1 = CreateObject("vbscript.regexp")
    With re1
       .Global = True
       .Pattern = "(\d{4})[\/|-](\d{1,2})[\/|-](\d{1,2})(\s\w+|)\s(\d{1,2}):(\d{1,2})(:\d{1,2}|)"
    End With
    
    Set re2 = CreateObject("vbscript.regexp")
    With re2
       .Global = True
       .Pattern = "(\d{4})[\/|-](\d{1,2})[\/|-](\d{1,2})(\s\w+|)\s(\d{1,2}):(\d{1,2})(:\d{1,2}|)"
    End With
    
End Function

Function decRaw()
    Dim id As String
    
    If InStr(1, Cells(1, 1).Value, "SDR", 1) > 0 Then
        id = decInfoSDR
        
        decDataSDR id
    ElseIf InStr(1, Cells(1, 1).Value, "Multi-Track Export -", 1) > 0 Then
        id = decInfoNomad
    End If
    
    
End Function


' SDR 格式信息提取

Function decInfoSDR() As String
    Dim s As New Station
    Dim ss As Sensor
    
    s.System = "SDR"
    s.Version = Cells(1, 2).Value
    
    Dim sensors(1 To MaxSensorNum) As Sensor
    Dim ssIndex As Integer
    ssIndex = 1
    
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
            End With
            
            If Len(ss.Channel) < 1 Then
                Call Error("传感器Channel号为空")
            End If
            
            Select Case ss.Units
                Case "", "-----", "unit"
                    ss.NotInstalled = True
                Case Else
                    ss.NotInstalled = False
            End Select
            
            Set myMatches = reISH.Execute(Cells(i + 5, 2).Value)
            If myMatches.Count >= 1 Then
                Set myMatch = myMatches(0)
                If myMatch.SubMatches.Count >= 2 Then
                    ss.Height = myMatch.SubMatches(0)
        
                    If myMatch.SubMatches(1) = "ft" Then
                        ss.Height = ss.Height * 0.3048
                    End If
                End If
            End If
            
            s.SensorsR.add CStr(i), ss
            ssIndex = ssIndex + 1
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

Function getSensor(i As Single) As Sensor
    Dim ss As New Sensor
    With ss
        .Channel = Cells(i, 2).Value
        .Cat = Cells(i + 1, 2).Value
        .Description = Cells(i + 2, 2).Value
        .Details = Cells(i + 3, 2).Value
        .SerialNumber = Cells(i + 4, 2).Value
        .ScaleFactor = Cells(i + 6, 2).Value
        .Offset = Cells(i + 7, 2).Value
        .Units = Cells(i + 8, 2).Value
    End With
    
    If Len(ss.Channel) < 1 Then
        Call Error("传感器Channel号为空")
    End If
    
    Select Case ss.Units
        Case "", "-----", "unit"
            ss.NotInstalled = True
        Case Else
            ss.NotInstalled = False
    End Select
    
    Set myMatches = reISH.Execute(Cells(i + 5, 2).Value)
    If myMatches.Count >= 1 Then
        Set myMatch = myMatches(0)
        If myMatch.SubMatches.Count >= 2 Then
            ss.Height = myMatch.SubMatches(0)

            If myMatch.SubMatches(1) = "ft" Then
                ss.Height = ss.Height * 0.3048
            End If
        End If
    End If
    
    Set getSensor = ss
End Function

Function addStation(s As Station)

    If Stations.Count >= 1 Then
        If Stations.Exists(s.id) Then
            Info "站点已存在: " + CStr(s.id)
            Exit Function
        End If
    End If
    
    Info "新增站点: " + CStr(s.id)
    Stations.add s.id, s
    
    showInfo s
    
End Function

Function showInfo(s As Station)
    Dim fs As Object
    Set fs = ActiveSheet
    Dim maxH As Single
    
    Sheets.add After:=Sheets(WB.Sheets.Count)
    
    Columns("A:A").HorizontalAlignment = xlCenter
    Columns("B:B").HorizontalAlignment = xlCenter
    Columns("C:C").HorizontalAlignment = xlCenter
    

    Range("A1:C1").Merge
    Range("A1:C1").Value = s.id + "测风塔配置一览表"
    
    Range("A2").Value = "测风塔"
    Range("B2:C2").Merge
    Range("B2:C2").Value = s.id
    
    Range("A3").Value = "地理位置" ' 必要时需要进行转换
    Range("B3:C3").Merge
    Range("B3:C3").Value = s.Site.Latitude + "," + s.Site.Longitude

    Range("A4").Value = "海拔高度"
    Range("B4:C4").Merge
    Range("B4:C4").Value = CStr(s.Site.SiteElevation) + " m"
    
    Range("A5").Value = "测风时段"
    Range("B5:C5").Merge
    Range("B5:C5").Value = "start～end" ' 在生成data后写入
    
    Range("A6").Value = "塔高"
    Range("B6:C6").Merge
    Range("B6:C6").Value = "70 m" ' 求取height最大值
    
    ' sensor 分类
    Range("A7").Value = "信道"
    Range("B7").Value = "安装高度 (m)"
    Range("C7").Value = "观测项目"
    
    Dim wv As New Scripting.Dictionary
    Dim wd As New Scripting.Dictionary
    Dim p As New Scripting.Dictionary
    Dim t As New Scripting.Dictionary
    Dim h As New Scripting.Dictionary
    Dim vol As New Scripting.Dictionary

    Dim ss As Sensor
    Dim i As Integer
    i = 8
    For Each key In s.SensorsR.Keys
        Set ss = s.SensorsR(key)
        
        If ss.Height > maxH Then
            maxH = ss.Height
        End If
        
        Select Case ss.Units
            Case "m/s", "mph"
                wv.add ss.Channel, ss

                addInfoSensor i, ss.Channel, ss.Height, "风速 (m/s)"
                i = i + 1
            Case "deg", "Degrees"
                wd.add ss.Channel, ss
                addInfoSensor i, ss.Channel, ss.Height, "风向 (度)"
                i = i + 1
            Case "Volts", "v"
                vol.add ss.Channel, ss
            Case "%RH"
                h.add ss.Channel, ss
            Case "C", "Degrees F"
                t.add ss.Channel, ss
                addInfoSensor i, ss.Channel, ss.Height, "气温 (℃)"
                i = i + 1
            Case "kPa", "mb", "mB"
                p.add ss.Channel, ss
                addInfoSensor i, ss.Channel, ss.Height, "气压 (kpa)"
                i = i + 1
        End Select
        
    Next
    
    Range("B6:C6").Value = CStr(maxH) + " m"
    
    Columns("A:A").EntireColumn.AutoFit
    Columns("B:B").ColumnWidth = 16
    Columns("C:C").ColumnWidth = 15
    
    ActiveSheet.Name = "info-" + s.id

    fs.Select

End Function

Function addInfoSensor(i As Integer, c As String, h As Single, t As String)
    Range("A" + CStr(i)).Value = "CH" + c
    Range("B" + CStr(i)).Value = h
    Range("C" + CStr(i)).Value = t
End Function

Function decDataSDR(id As String)
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
    
    ' 调整日期格式
    
    Dim i
    For i = 2 To ActiveSheet.UsedRange.Rows.Count
        Cells(i, 1).Value = decDate(Cells(i, 1).Value)
    Next i
    

    
    maxX = ActiveSheet.UsedRange.Rows.Count
    maxY = ActiveSheet.UsedRange.Columns.Count
    x = Cells(2, maxY + 1).Address
    y = Cells(maxX, maxY + 1).Address
    
    Cells(2, maxY + 1).Formula = "=MINUTE(A2)"
    Cells(2, maxY + 1).AutoFill Destination:=Range(x + ":" + y)

    Cells(1, maxY + 1).Formula = "=AVERAGE(" + CStr(x) + ":" + CStr(y) + ")"
    
    If Cells(1, maxY + 1).Value > 1 Then
        ActiveSheet.Name = "data-" + id + "-10m"
    Else
        ActiveSheet.Name = "data-" + id + "-1h"
    End If
    
    
End Function


Function decDate(str As String) As String
    Set myMatches = re1.Execute(str)
    If myMatches.Count >= 1 Then
        Set myMatch = myMatches(0)
        If myMatch.SubMatches.Count >= 6 Then
            decDate = newDate(myMatch.SubMatches(0), myMatch.SubMatches(1), myMatch.SubMatches(2), myMatch.SubMatches(4), myMatch.SubMatches(5))
            Exit Function
        End If
    End If
    
    Set myMatches = re2.Execute(str)
    If myMatches.Count >= 1 Then
        Set myMatch = myMatches(0)
        If myMatch.SubMatches.Count >= 6 Then
            decDate = newDate(myMatch.SubMatches(2), myMatch.SubMatches(0), myMatch.SubMatches(1), myMatch.SubMatches(3), myMatch.SubMatches(4))
            Exit Function
        End If
    End If
    
    Error "时间格式错误"
    
End Function

Function newDate(y As Integer, mo As Integer, d As Integer, h As Integer, min As Integer) As String
    newDate = CStr(y) + "/" + CStr(mo) + "/" + CStr(d) + " " + CStr(h) + ":" + CStr(min)
End Function


' Nomad 格式信息提取

Function decInfoNomad()
    Dim s As Station
    s.System = "Nomad"
    


End Function


