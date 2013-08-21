Attribute VB_Name = "读取"

Public re1      As Object   ' 判断时间正则1:
Public re2      As Object   ' 判断时间正则2:
Public re3      As Object

Sub 读取数据()
    系统初始化
    init
    
    Dim st As Object
    
    For Each st In Sheets
        If InStr(1, st.Name, "raw", 1) > 0 Then
            decRaw st
        End If
    Next

    For Each k In Stations
        Dim s As Object: Set s = Stations(k)
        
        sensorClassfy s
        
        deleteZero stn:=s.Sheet1h, s:=s
        deleteZero stn:=s.Sheet10m, s:=s
        
    Next
End Sub
Private Function init()
    ' 设定正则
    
    Set re1 = CreateObject("vbscript.regexp")
    re1.Pattern = "(\d{4})[\/|-](\d{1,2})[\/|-](\d{1,2})(\s\w+|)\s(\d{1,2}):(\d{1,2})(:\d{1,2}|)"
    ' yyyy/m/d (weekday) h:mm
    ' 2011/11/18 Friday 07:10:00
    ' 2011/11/18 07:10:00

    Set re2 = CreateObject("vbscript.regexp")
    re2.Pattern = "(\d{1,2})[\/|-](\d{1,2})[\/|-](\d{4})\s(\d{1,2}):(\d{1,2})(:\d{1,2}|)"
    ' m/d/yyyy h:mm
    ' 1/2/2010 0:00

    Set re3 = CreateObject("vbscript.regexp")
    re3.Pattern = "^""?(\d{4})[\-|\/](\d{1,2})[\-|\/](\d{1,2})""?$"
    ' yyyy/m/d "yyyy/m/d"
    ' 2012/1/1
    ' "2012/1/1"
    
    ' yyyy/d/m h:mm 0004-1-10m.txt
    ' 2012/1/1 23:50
    ' 2012/2/1 0:00
End Function
Private Function deleteZero(stn As String, s As Object)
    If stn <> "" And sheetExist(stn) Then
        Dim st As Object: Set st = Sheets(stn)
        
        st.UsedRange.AutoFilter
        
        Dim wvs As Object: Set wvs = s.sensors("风速")
        Dim a: a = wvs.Items
        For j = 0 To wvs.count - 1
            Dim ss As Object: Set ss = a(j)
            
            st.UsedRange.AutoFilter Field:=ss.avg, Criteria1:="<>0", Operator:=xlAnd
        Next
        
        Dim tt As Object: Set tt = newSheet("tdeletezero")
        rangeCopy st.UsedRange, tt.Cells(1, 1)
        Dim n As String: n = st.Name
        deleteSheet st
        tt.Name = n
        tt.Columns(1).NumberFormatLocal = "yyyy/m/d h:mm"
    End If
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
    If Stations.count >= 1 Then
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
    If mymatches.count >= 1 Then
        Set mymatch = mymatches(0)
        If mymatch.SubMatches.count >= 6 Then
            decDate = newDate(mymatch.SubMatches(0), mymatch.SubMatches(1), mymatch.SubMatches(2), mymatch.SubMatches(4), mymatch.SubMatches(5))
            Exit Function
        End If
    End If
    
    Set mymatches = re2.Execute(str)
    If mymatches.count >= 1 Then
        Set mymatch = mymatches(0)
        If mymatch.SubMatches.count >= 6 Then
            decDate = newDate(mymatch.SubMatches(2), mymatch.SubMatches(0), mymatch.SubMatches(1), mymatch.SubMatches(3), mymatch.SubMatches(4))
            Exit Function
        End If
    End If
    
    Set mymatches = re3.Execute(str)
    If mymatches.count >= 1 Then
        Set mymatch = mymatches(0)
        If mymatch.SubMatches.count >= 3 Then
            decDate = newDate(mymatch.SubMatches(0), mymatch.SubMatches(1), mymatch.SubMatches(2), 0, 0)
            Exit Function
        End If
    End If
    
    decDate = str
    Err "时间格式错误: " + str
    
End Function

Function newDate(y As Integer, mo As Integer, d As Integer, h As Integer, Min As Integer) As String
    newDate = CStr(y) + "/" + CStr(mo) + "/" + CStr(d) + " " + CStr(h) + ":" + CStr(Min)
End Function

Function adjustData(ds As Object, s As Object)
    ' 调整日期格式

    For i = 2 To ds.UsedRange.Rows.count
        ds.Cells(i, 1).Value = decDate(ds.Cells(i, 1).Value)
    Next i
    
    ds.Columns(1).NumberFormatLocal = "yyyy/m/d h:mm"
    
    ' 判断是10分钟还是60分钟数据
    
    Dim maxX: maxX = ds.UsedRange.Rows.count
    Dim maxY: maxY = ds.UsedRange.Columns.count
    
    Dim t As Object: Set t = newSheet("tadjustdata")
    t.Cells(2, 1).Formula = "=MINUTE('" + ds.Name + "'!A2)"
    t.Cells(2, 1).AutoFill Destination:=t.Range(t.Cells(2, 1), t.Cells(maxX, 1))
    
    Dim avg As Double: avg = Application.WorksheetFunction.Average(t.Columns(1))
    deleteSheet t

    Dim n As String
    If avg > 1 Then
        n = "data" + s.id + "10m"
        s.Sheet10m = n
    Else
        n = "data" + s.id + "1h"
        s.Sheet1h = n
    End If
    If sheetExist(n) Then
        deleteSheet Sheets(n)
    End If
    ds.Name = n
    ds.Range(ds.Cells(2, 2), ds.Cells(ds.UsedRange.Rows.count, ds.UsedRange.Columns.count)).NumberFormatLocal = "0.0"

    ' 起始结束时间
    s.StartTime = Format(Application.WorksheetFunction.Min(ds.Columns(1)), "yyyy/m/d h:mm")
    s.EndTime = Format(Application.WorksheetFunction.max(ds.Columns(1)), "yyyy/m/d h:mm")

End Function

Sub 生成1h()
    系统初始化
    
    Dim k, st As Object
    For Each k In Stations
        Set s = Stations(k)
        
        If s.Sheet1h = "" Then
            If s.Sheet10m = "" Then
                MsgBox s.id + ": no 10m data!"
                Exit Sub
            End If
            Info s.id + " no 1h data! gen from 10m data."
            
            Dim d1h As Object: Set d1h = newSheet("data" + s.id + "1h")
            genD1fD2 Sheets(s.Sheet10m), d1h
            s.Sheet1h = "data" + s.id + "1h"
        End If
    Next
End Sub
Private Function genD1fD2(d2 As Object, d1 As Object)
    Dim maxX: maxX = d2.UsedRange.Rows.count
    Dim maxY: maxY = d2.UsedRange.Columns.count

    rangeCopy d2.Rows(1), d1.Cells(1, 1)
    
    Dim d2s, d2e, d1i, val
    For i = 2 To maxX
        If i = 2 Then
            d2s = 2
            d2e = 2
            d1i = 2
            val = formatYMDH(d2.Cells(i, 1).Value)
        End If
    
        If formatYMDH(d2.Cells(i, 1).Value) = val Then
            d2e = d2e + 1
        Else
            d1.Cells(d1i, 1).Value = val

            For j = 2 To maxY
                If d2.Cells(d2s, j).Value <> "" Then
                    d1.Cells(d1i, j).Value = Application.WorksheetFunction.Average(Range(d2.Cells(d2s, j), d2.Cells(d2e, j)))
                End If
            Next j

            d2s = i
            d2e = i
            d1i = d1i + 1
            val = formatYMDH(d2.Cells(i, 1).Value)
        End If
        
    Next i
    
    d1.Range(d1.Cells(2, 2), d1.Cells(d1.UsedRange.Rows.count, d1.UsedRange.Columns.count)).NumberFormatLocal = "0.0"
End Function
Private Function formatYMDH(v As Double) As String
    formatYMDH = Format(v, "yyyy/mm/dd hh:00")
End Function


Function sensorClassfy(s As Object)
    Dim k, ss As Object
    For Each k In s.sensorsR
        Set ss = s.sensorsR(k)
        Select Case ss.Units
            Case "m/s", "mph"

                If ss.Units = "mph" Then
                    adjustRTimes s.Sheet10m, ss, 1.6 / 3.6
                    adjustRTimes s.Sheet1h, ss, 1.6 / 3.6
                    ss.Units = "m/s"
                End If
                
                ss.Scat = "风速" '"wv"
            Case "deg", "Degress", "Degrees F", "F"

                If ss.Units = "Degrees F" Or ss.Units = "F" Then
                    adjustRF s.Sheet10m, ss
                    adjustRF s.Sheet1h, ss
                End If
                ss.Units = "deg"
                
                ss.Scat = "风向" '"wd"
            Case "Volts", "v"
                ss.Scat = "vol"
            Case "%RH"
                ss.Scat = "h"
            Case "C", "Degrees F", "F"
                ss.Scat = "气温" '"t"
            Case "kPa", "mb", "mB", "MB"

                If ss.Units = "mb" Or ss.Units = "mB" Or ss.Units = "MB" Then
                    adjustRTimes s.Sheet10m, ss, 0.1
                    adjustRTimes s.Sheet1h, ss, 0.1
                    ss.Units = "kPa"
                End If
                
                ss.Scat = "气压" '"p"
        End Select
    Next
End Function
Private Function adjustRTimes(sn As String, ss As Object, t As Double)
    If sn <> "" And sheetExist(sn) Then
        Dim r As Object: Set r = getAdjustRange(sn, ss)
        
        rangeFV r, t, xlMultiply
        
        r.NumberFormatLocal = "0.0"
    End If
End Function
Private Function adjustRF(sn As String, ss As Object)
    If sn <> "" And sheetExist(sn) Then
        Dim r As Object: Set r = getAdjustRange(sn, ss)

        rangeFV r, 32, xlSubtract
        rangeFV r, 1.8, xlDivide
        
        r.NumberFormatLocal = "0.0"
    End If
End Function
Private Function getAdjustRange(sn As String, ss As Object) As Object
    Dim st As Object: Set st = Sheets(sn)
    Dim maxX As Integer: maxX = st.UsedRange.Rows.count
    Dim r As Object: Set r = st.Range(st.Cells(2, ss.avg), st.Cells(maxX, ss.avg))
    
    If ss.max > 0 Then
        Set r = Application.Union(r, st.Range(st.Cells(2, ss.max), st.Cells(maxX, ss.max)))
    End If
    
    If ss.Min > 0 Then
        Set r = Application.Union(r, st.Range(st.Cells(2, ss.Min), st.Cells(maxX, ss.Min)))
    End If
    
    Set getAdjustRange = r
End Function

'Latitude    N 038?11.591'
'Latitude    S 028? 54.473'
'Latitude    N 39?16.336?
'Latitude    N 029 36.282'
'Latitude    N 000?00.000'
'Latitude    N 000 00.000'
'Latitude
'Latitude    0S
'Latitude    374027N
Function adjustLL(str As String) As String
    Dim reNSS As Object: Set reNSS = CreateObject("vbscript.regexp")
    reNSS.Pattern = "(\d*)(N|W|S|E)\s*(\d*)[\?\s]*([\d\.]*)['\?]*"
    
    Set mymatches = reNSS.Execute(str)
    If mymatches.count > 0 Then
        Set mymatch = mymatches(0).SubMatches

        If mymatch.count >= 1 Then
            If mymatch(0) <> "" Then
                adjustLL = "not match:" + str
            
            Else
                adjustLL = mymatch(1) + " " + mymatch(2) + "°" + mymatch(3) + "'"
            
            End If
        End If
    Else
        adjustLL = "not match:" + str
    End If

End Function

