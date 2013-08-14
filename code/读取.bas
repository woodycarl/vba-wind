Attribute VB_Name = "读取"




Public re1      As Object   ' 判断时间正则1:
Public re2      As Object   ' 判断时间正则2:
Public re3      As Object

Sub 读取数据()
    系统初始化
    initRead
    
    Dim st As Object
    
    For Each st In Sheets
        If InStr(1, st.Name, "raw", 1) > 0 Then
            decRaw st
        End If
    Next
    
    Dim k, s As Object
    For Each k In Stations
        Set s = Stations(k)
        
        sensorClassfy s
    Next
End Sub

Sub 生成1h()

    Dim k, st As Object
    For Each k In Stations
        Set st = Stations(k)
        
        If st.Sheet1h = "" Then
            If st.Sheet10m = "" Then
                MsgBox st.id + ": no data!"
                Exit Sub
            End If
            Info st.id + " no 1h data! gen from 10m data."
            
            Sheets.Add After:=Sheets(Sheets.Count)
            ActiveSheet.Name = "data" + st.id + "1h"
            
            genD1fD2 Sheets(st.Sheet10m), ActiveSheet
            
            st.Sheet1h = "data" + st.id + "1h"
        End If
    Next
End Sub

' 设定正则
Function initRead()


    Set re1 = CreateObject("vbscript.regexp")
    re1.Pattern = "(\d{4})[\/|-](\d{1,2})[\/|-](\d{1,2})(\s\w+|)\s(\d{1,2}):(\d{1,2})(:\d{1,2}|)"

    Set re2 = CreateObject("vbscript.regexp")
    re2.Pattern = "(\d{1,2})[\/|-](\d{1,2})[\/|-](\d{4})\s(\d{1,2}):(\d{1,2})(:\d{1,2}|)"
    
    Set re3 = CreateObject("vbscript.regexp")
    re3.Pattern = "^""?(\d{4})[\-|\/](\d{1,2})[\-|\/](\d{1,2})""?$"
End Function

Function decRaw(st As Object)
    Dim v: v = st.Range("A1").Value
    
    If InStr(1, v, "SDR", 1) > 0 Then
        decSDR st
    ElseIf InStr(1, v, "Multi-Track Export -", 1) > 0 Then
        decNomad st
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
    Stations.Add s.id, s
        
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
    
    Set mymatches = re3.Execute(str)
    If mymatches.Count >= 1 Then
        Set mymatch = mymatches(0)
        If mymatch.SubMatches.Count >= 3 Then
            decDate = newDate(mymatch.SubMatches(0), mymatch.SubMatches(1), mymatch.SubMatches(2), 0, 0)
            Exit Function
        End If
    End If
    
    MsgBox "time format err: " + str
    'Error "时间格式错误"
    
End Function

Function newDate(y As Integer, mo As Integer, d As Integer, h As Integer, Min As Integer) As String
    newDate = CStr(y) + "/" + CStr(mo) + "/" + CStr(d) + " " + CStr(h) + ":" + CStr(Min)
End Function

Function adjustData(ds As Object, s As Object)
    ' 调整日期格式
    
    Dim i
    For i = 2 To ds.UsedRange.Rows.Count
        ds.Cells(i, 1).Value = decDate(ds.Cells(i, 1).Value)
    Next i
    
    ds.Columns("A:A").NumberFormatLocal = "yyyy/m/d h:mm"
    
    ' 判断是10分钟还是60分钟数据
    
    Dim maxX, maxY
    maxX = ds.UsedRange.Rows.Count
    maxY = ds.UsedRange.Columns.Count
    
    Dim x As String, y As String
    x = ds.Cells(2, maxY + 1).Address
    y = ds.Cells(maxX, maxY + 1).Address
    
    ds.Cells(2, maxY + 1).Formula = "=MINUTE(A2)"
    ds.Cells(2, maxY + 1).AutoFill Destination:=ds.Range(x + ":" + y)

    ds.Cells(1, maxY + 1).Formula = "=AVERAGE(" + CStr(x) + ":" + CStr(y) + ")"
    
    Dim n As String
    If ds.Cells(1, maxY + 1).Value > 1 Then
        n = "data" + s.id + "10m"
        s.Sheet10m = n
    Else
        n = "data" + s.id + "1h"
        s.Sheet1h = n
    End If
    If sheetExist(n) Then
        Sheets(n).Delete
    End If
    ds.Name = n
    
    ds.Range(ds.Cells(1, maxY + 1).Address + ":" + ds.Cells(maxX, maxY + 1).Address).Clear
    
    ' 起始结束时间
    
    ds.Cells(2, maxY + 1).Formula = "=min('" + ds.Name + "'!A:A)"
    ds.Cells(2, maxY + 1).NumberFormatLocal = "yyyy/m/d h:mm"
    s.StartTime = ds.Cells(2, maxY + 1).Value
    ds.Cells(2, maxY + 1).Formula = "=max('" + ds.Name + "'!A:A)"
    ds.Cells(2, maxY + 1).NumberFormatLocal = "yyyy/m/d h:mm"
    s.EndTime = ds.Cells(2, maxY + 1).Value
    
    ds.Cells(2, maxY + 1).Clear
End Function


Function genD1fD2(d2 As Object, d1 As Object)
    Dim maxX, maxY
    maxY = d2.UsedRange.Columns.Count
    maxX = d2.UsedRange.Rows.Count

    d2.Range(d2.Cells(1, 1), d2.Cells(1, maxY)).Copy
    d1.Paste
    
    Dim d2s, d2e, d1i, val
    
    Dim i, j, k
    Dim sum
    For i = 2 To maxX
        If i = 2 Then
            d2s = 2
            d2e = 2
            d1i = 2
            val = Format(d2.Cells(i, 1).Value, "yyyy/mm/dd hh")
        End If
    
        If Format(d2.Cells(i, 1).Value, "yyyy/mm/dd hh") = val Then
            d2e = d2e + 1
        Else
            d1.Cells(d1i, 1).Value = val + ":00"
            
            
            For j = 2 To maxY

                If d2.Cells(d2s, j).Value <> "" Then
                    
                    sum = 0
                    For k = d2s To d2e
                        sum = sum + d2.Cells(k, j).Value
                    Next k
                    d1.Cells(d1i, j) = sum / (d2e - d2s + 1)
                    'd1.Cells(d1i, j).Formula = "=average(" + d2.Name + "!" + d2.Cells(d2s, j).Address + ":" + d2.Cells(d2e, j).Address + ")"
                End If
            Next j
            
            d2s = i
            d2e = i
            d1i = d1i + 1
            val = Format(d2.Cells(i, 1).Value, "yyyy/mm/dd hh")
        End If
        
    Next i
    
    d1.Range(Cells(2, 2), Cells(d1.UsedRange.Rows.Count, d1.UsedRange.Columns.Count)).NumberFormatLocal = "0.00_);[红色](0.00)"

End Function




Function adjustRTimes(sn As String, ss As Object, t As Double)
    Dim i, ds As Object
    
    If sn <> "" And sheetExist(sn) Then

        Set ds = Sheets(sn)
        For i = 2 To ds.Rows.Count
            If ss.avg > 0 Then
                ds.Cells(i, ss.avg).Value = ds.Cells(i, ss.avg).Value * t
            End If

            If ss.Max > 0 Then
                ds.Cells(i, ss.Max).Value = ds.Cells(i, ss.Max).Value * t
            End If
            
            If ss.Min > 0 Then
                ds.Cells(i, ss.Min).Value = ds.Cells(i, ss.Min).Value * t
            End If
        Next i
    End If
End Function

Function adjustRF(sn As String, ss As Object)
    Dim i, ds As Object
    
    If sn <> "" And sheetExist(sn) Then
        Set ds = Sheets(sn)
        For i = 2 To ds.Rows.Count
            If ss.avg > 0 Then
                ds.Cells(i, ss.avg).Value = (ds.Cells(i, ss.avg).Value - 32) / 1.8
            End If
            
            If ss.Max > 0 Then
                ds.Cells(i, ss.Max).Value = (ds.Cells(i, ss.Max).Value - 32) / 1.8
            End If
            
            If ss.Min > 0 Then
                ds.Cells(i, ss.Min).Value = (ds.Cells(i, ss.Min).Value - 32) / 1.8
            End If
        Next i
    End If
End Function


Function sensorClassfy(s As Object)
    Dim k, ss As Object
    For Each k In s.SensorsR
        Set ss = s.SensorsR(k)
        Select Case ss.Units
            Case "m/s", "mph"

                If ss.Units = "mph" Then
                    adjustRTimes s.Sheet10m, ss, 1.6 / 3.6
                    adjustRTimes s.Sheet1h, ss, 1.6 / 3.6
                    ss.Units = "m/s"
                End If
                
                ss.Scat = "wv"
            Case "deg", "Degress", "Degrees F", "F"

                If ss.Units = "Degrees F" Or ss.Units = "F" Then
                    adjustRF s.Sheet10m, ss
                    adjustRF s.Sheet1h, ss
                End If
                ss.Units = "deg"
                
                ss.Scat = "wd"
            Case "Volts", "v"
                ss.Scat = "vol"
            Case "%RH"
                ss.Scat = "h"
            Case "C", "Degrees F", "F"
                ss.Scat = "t"
            Case "kPa", "mb", "mB", "MB"

                If ss.Units = "mb" Or ss.Units = "mB" Or ss.Units = "MB" Then
                    adjustRTimes s.Sheet10m, ss, 0.1
                    adjustRTimes s.Sheet1h, ss, 0.1
                    ss.Units = "kPa"
                End If
                
                ss.Scat = "p"
        End Select
    Next
End Function

