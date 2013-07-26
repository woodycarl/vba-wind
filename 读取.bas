Attribute VB_Name = "读取"


Public reISH    As Object   ' info sensor height
Public re1      As Object   ' 判断时间正则1:
Public re2      As Object   ' 判断时间正则2:


Sub 读取数据()
    initRead
    
    Dim s, fs As Object
    Set fs = ActiveSheet

    For Each s In Sheets
        If InStr(1, s.Name, "raw", 1) > 0 Then
            s.Select

            decRaw
        End If
    Next

    fs.Select
    
End Sub

' 设定正则
Function initRead()
    Set reISH = CreateObject("vbscript.regexp")
    reISH.Pattern = "^([\d\.]+)\s*(m|ft)"

    Set re1 = CreateObject("vbscript.regexp")
    re1.Pattern = "(\d{4})[\/|-](\d{1,2})[\/|-](\d{1,2})(\s\w+|)\s(\d{1,2}):(\d{1,2})(:\d{1,2}|)"

    Set re2 = CreateObject("vbscript.regexp")
    re2.Pattern = "(\d{1,2})[\/|-](\d{1,2})[\/|-](\d{4})\s(\d{1,2}):(\d{1,2})(:\d{1,2}|)"
End Function

Function decRaw()
    Dim v
    v = Range("A1").Value
    
    If InStr(1, v, "SDR", 1) > 0 Then
        decDataSDR decInfoSDR
    ElseIf InStr(1, v, "Multi-Track Export -", 1) > 0 Then
        decDataNomad decInfoNomad
    End If

End Function

Function addStation(s As Station)

    If Stations.Count >= 1 Then
        If Stations.Exists(s.id) Then
            Info "站点已存在: " + CStr(s.id)
            Exit Function
        End If
    End If
    
    Info "新增站点: " + s.id
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

Function decDate(str As String) As String
    Set mymatches = re1.Execute(str)
    If mymatches.Count >= 1 Then
        Set mymatch = mymatches(0)
        If mymatch.SubMatches.Count >= 6 Then
            decDate = newDate(mymatch.SubMatches(0), mymatch.SubMatches(1), mymatch.SubMatches(2), mymatch.SubMatches(4), mymatch.SubMatches(5))
            Exit Function
        End If
    End If
    
    Set mymatches = re2.Execute(str)
    If mymatches.Count >= 1 Then
        Set mymatch = mymatches(0)
        If mymatch.SubMatches.Count >= 6 Then
            decDate = newDate(mymatch.SubMatches(2), mymatch.SubMatches(0), mymatch.SubMatches(1), mymatch.SubMatches(3), mymatch.SubMatches(4))
            Exit Function
        End If
    End If
    
    Error "时间格式错误"
    
End Function

Function newDate(y As Integer, mo As Integer, d As Integer, h As Integer, Min As Integer) As String
    newDate = CStr(y) + "/" + CStr(mo) + "/" + CStr(d) + " " + CStr(h) + ":" + CStr(Min)
End Function

Function adjustData(id As String)
    ' 需要激活sheet
    ''''''''''''''''

    ' 调整日期格式
    
    Dim i
    For i = 2 To ActiveSheet.UsedRange.Rows.Count
        Cells(i, 1).Value = decDate(Cells(i, 1).Value)
    Next i
    
    ' 判断是10分钟还是60分钟数据
    
    Dim maxX, maxY
    maxX = ActiveSheet.UsedRange.Rows.Count
    maxY = ActiveSheet.UsedRange.Columns.Count
    
    Dim x As String, y As String
    x = Cells(2, maxY + 1).Address
    y = Cells(maxX, maxY + 1).Address
    
    Cells(2, maxY + 1).Formula = "=MINUTE(A2)"
    Cells(2, maxY + 1).AutoFill Destination:=Range(x + ":" + y)

    Cells(1, maxY + 1).Formula = "=AVERAGE(" + CStr(x) + ":" + CStr(y) + ")"
    
    Dim n As String
    If Cells(1, maxY + 1).Value > 1 Then
        n = "data-" + id + "-10m"
    Else
        n = "data-" + id + "-1h"
    End If
    If sheetExist(n) Then
        Sheets(n).Delete
    End If
    ActiveSheet.Name = n
    
    Range(Cells(1, maxY + 1).Address + ":" + Cells(maxX, maxY + 1).Address).Clear
End Function



