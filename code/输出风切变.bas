Attribute VB_Name = "输出风切变"

Function 计算风切变指数(s As Object, rst As Object, dst As Object)
        ' 代表年的不同高度风切变指数
        s.Pc.Value = "代表年的不同高度风切变指数"
        s.Pc = s.Pc.Offset(1, 0)
        
        Dim wvs As Object: Set wvs = s.Sensors("wv")
        Dim a: a = wvs.Items
        
        Dim wss As New Collection
        For j = 0 To wvs.count - 1
            Dim ss As Object: Set ss = a(j)
            If ss.height = 0 Then
                Err s.id + "计算风切变指数: CH" + ss.channel + " 高度为空"
                GoTo cf1
            End If
            
            Dim twss As WS: Set twss = New WS
            With twss
                .height = ss.height
                .avg = Application.WorksheetFunction.Average(rst.Columns(ss.avg))
            End With
            wss.Add twss
cf1:
        Next j
        
        If wss.count < 2 Then
            Err s.id + "计算风切变指数: 可用参数不足!"
            Exit Function
        End If

        drawWS wss, dst, s.Pc
        s.Pc = s.Pc.Offset(wss.count + 18, 0)
End Function

Sub 绘制选定多列的风切变图()
    系统初始化
    Dim st As Object: Set st = ActiveSheet
    Dim ar: ar = Split(Selection.Address, ",")
    
    Dim wss As New Collection

    Dim cArr As New Collection
    
    For i = 0 To UBound(ar)
        Dim tr As Object: Set tr = st.Range(ar(i))

        If tr.Rows.count >= tr.Columns.count Then
            For j = 1 To tr.Columns.count
                cArr.Add tr.Columns(j)
            Next j
        ElseIf tr.Columns.count > tr.Rows.count Then
            For j = 1 To tr.Rows.count
                cArr.Add tr.Rows(j)
            Next j
        End If
    Next i
    
    For Each c In cArr
        Dim th As Double: th = InputBox("输入序列 " + c.Address + "的高度:")
        Dim avg As Double: avg = Application.WorksheetFunction.Average(c)
        Dim twss As WS: Set twss = New WS
        With twss
            .height = th
            .avg = avg
        End With
        wss.Add twss
    Next
    
    Dim wbn As Object: Set wbn = Workbooks.Add
    Dim dst As Object: Set dst = wbn.Sheets(1)
    Dim dr As Object: Set dr = dst.Range("A1")
    
    drawWS wss, dst, dr
End Sub

Function drawWS(wss As Collection, dst As Object, dr As Object)
    dr.Value = "高度"
    dr.Offset(0, 1).Value = "风速 (m/s)"
    
    For i = 1 To wss.count
        dr.Offset(i, 0).Value = wss(i).height
        dr.Offset(i, 1).Value = wss(i).avg
    Next
    
    Dim maxX: maxX = dst.UsedRange.Rows.count

    Dim rangeH: Set rangeH = dst.Range(dr.Offset(1, 0), dr.Offset(wss.count, 0))
    Dim rangeA: Set rangeA = dst.Range(dr.Offset(1, 1), dr.Offset(wss.count, 1))

    Dim rangeSort: Set rangeSort = dst.Range(dr.Offset(0, 0).Address + ":" + dr.Offset(wss.count, 1).Address)

    dst.Sort.SortFields.Add Key:=rangeH, _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With dst.Sort
        .SetRange rangeSort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    For i = 1 To wss.count
        dr.Offset(0, 1 + i).Value = dr.Offset(i, 0).Value
        
        dr.Offset(i + wss.count, 0).Value = log(dr.Offset(i, 0).Value)
        dr.Offset(i + wss.count, 1).Value = log(dr.Offset(i, 1).Value)
        
        For j = 1 To wss.count
            If dr.Offset(i, 0).Value > dr.Offset(j, 0).Value Then
                dr.Offset(i, j + 1).Value = windshear(dr.Offset(i, 1).Value, dr.Offset(j, 1).Value, dr.Offset(i, 0).Value, dr.Offset(j, 0).Value)
            End If
        Next j
    Next i
    
    Dim range1 As Object: Set range1 = dst.Range(dr.Offset(1 + wss.count, 1).Address + ":" + dr.Offset(2 * wss.count, 1).Address)
    Dim range2 As Object: Set range2 = dst.Range(dr.Offset(1 + wss.count, 0).Address + ":" + dr.Offset(2 * wss.count, 0).Address)
    Rsq = Application.WorksheetFunction.Rsq(range1, range2)
    Slope = Application.WorksheetFunction.Slope(range1, range2)
    Intercept = Application.WorksheetFunction.Intercept(range1, range2)
    
    Dim a: a = Exp(Intercept)
    Dim b: b = Slope
    
    For i = 1 To wss.count
        dr.Offset(i, 2 + wss.count).Value = a * dr.Offset(i, 0).Value ^ b
    Next i
    
    Dim rangeT: Set rangeT = dst.Range(dr.Offset(1, 2 + wss.count), dr.Offset(wss.count, 2 + wss.count))
    
    Dim rangeX As String: rangeX = dst.Name + "!" + rangeH.Address
    Dim cRangeY As New Collection, cRangeT As New Collection

    cRangeY.Add rangeA.Address
    cRangeT.Add dst.Name + "!" + dr.Offset(0, 1).Address
    cRangeY.Add rangeT.Address
    cRangeT.Add dst.Name + "!" + dr.Offset(0, 2 + wss.count).Address
    
    Dim myChart As Object
    Set myChart = drawChart(rangeX:=rangeX, cRangeY:=cRangeY, cRangeT:=cRangeT, rst:=dst, dst:=dst, _
            dpo:=dr.Offset(1 + wss.count, 0), cLegend:=-1, axisFormatX:="0", axisTitleX:="高度 (m)", _
            axisTitleY:="风速 (m/s)", cType:=xlXYScatterSmoothNoMarkers)
    
    With myChart
        With .SeriesCollection(1) '
            .MarkerStyle = -4105
            .Format.Line.Visible = msoFalse
        End With
        
        .Axes(xlCategory).MinimumScale = dr.Offset(1, 0).Value - 5
    End With


    Dim tb As String: tb = "y = " & Format(a, "0.00") & "x" & Format(b, "0.00")
    Dim tbs As Integer: tbs = InStr(1, tb, "x", 1) + 1
    Dim tbo As Integer: tbo = Len(tb) - tbs + 1

    myChart.Shapes.AddTextbox(msoTextOrientationHorizontal, 417, 91.8, 103.8, _
        56.4).Select
    Dim otb As Object: Set otb = myChart.Shapes(1)
    otb.TextFrame.Characters.Text = tb & Chr(13) & _
                                          "R2 = " & Format(Rsq, "0.00")
    otb.TextFrame2.TextRange.Characters(tbs, tbo).Font. _
            BaselineOffset = 0.3
    otb.TextFrame2.TextRange.Characters(Len(tb) + 3, 1).Font. _
            BaselineOffset = 0.3

    dst.Range(dr.Offset(1, 1), dr.Offset(wss.count, wss.count + 2)).NumberFormatLocal = "0.00"

End Function


Sub 计算选定两列的风切变()
    Dim st As Object: Set st = ActiveSheet

    Dim d1r As Object, d2r As Object
    
    Dim ar: ar = Split(Selection.Address, ",")
    
    If UBound(ar) = 0 Then
        Dim ta: ta = largethan(Selection, 2)
        If ta = 1 Then
            Set d1r = Selection.Columns(1)
            Set d2r = Selection.Columns(2)
        ElseIf ta = 2 Then
            Set d1r = Selection.Rows(1)
            Set d2r = Selection.Rows(2)
        Else
            MsgBox "选定数据不符 " & Selection.Rows.count & " " & Selection.Columns.count
            Exit Sub
        End If
        
    ElseIf UBound(ar) = 1 Then
        Dim ta1: ta1 = largethan(st.Range(ar(0)), 1)
        Dim ta2: ta2 = largethan(st.Range(ar(1)), 1)
        Dim r1: Set r1 = st.Range(ar(0))
        Dim r2: Set r2 = st.Range(ar(1))
        
        If ta1 = 1 And ta2 = 1 And r1.Rows.count = r2.Rows.count Then
        ElseIf ta1 = 2 And ta2 = 2 And r1.Columns.count = r2.Columns.count Then
        Else
            MsgBox "选定数据不符 " & Selection.Rows.count & " " & Selection.Columns.count
            Exit Sub
        End If
        
        Set d1r = r1
        Set d2r = r2
    End If
    
    Dim h1 As Double, h2 As Double
    h1 = InputBox("输入序列1 " + d1r.Address + "高度:")
    h2 = InputBox("输入序列2 " + d2r.Address + "高度:")
    
    WS = windshearR(d1r, d2r, h1, h2)
    MsgBox "风切变值: " & WS
End Sub

Function windshearR(r1 As Object, r2 As Object, h1 As Double, h2 As Double) As Double
    windshearR = windshear(Application.WorksheetFunction.Average(r1), _
        Application.WorksheetFunction.Average(r2), h1, h2)
End Function

Function windshear(avg1 As Double, avg2 As Double, h1 As Double, h2 As Double) As Double
    windshear = log(avg2 / avg1) / log(h2 / h1)
End Function

Function largethan(o As Object, i As Integer) As Integer
    If o.Rows.count > i And o.Columns.count = i Then
        largethan = 1
    ElseIf o.Columns.count > i And o.Rows.count = i Then
        largethan = 2
    Else
        largethan = 0
    End If
End Function
