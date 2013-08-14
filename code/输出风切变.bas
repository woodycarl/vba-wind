Attribute VB_Name = "输出风切变"

Sub 计算风切变指数()
    系统初始化
    
    For Each k In Stations
        Dim s As Object: Set s = Stations(k)
        
        If s.CurRePo = "A1" Then
            initCalResult s
        End If
        
        Dim rst As Object: Set rst = Sheets(s.Sheet1h)
        Dim dst As Object: Set dst = Sheets(s.Result)


        ' 代表年的不同高度风切变指数
        Dim pc As Object: Set pc = dst.Range(s.CurRePo)
        pc.Value = "代表年的不同高度风切变指数"
        s.CurRePo = pc.Offset(1, 0).Address
        
        Dim wvs As Object: Set wvs = s.Sensors("wv")
        Dim a: a = wvs.Items
        
        Dim wss As New Collection
        For j = 0 To wvs.Count - 1
            Dim ss As Object: Set ss = a(j)
            
            'Dim th As Double: th = ss.height
            'Dim avg As Double: avg = Application.WorksheetFunction.Average(rst.Columns(ss.avg))
            Dim twss As WS: Set twss = New WS
            With twss
                .height = ss.height
                .avg = Application.WorksheetFunction.Average(rst.Columns(ss.avg))
            End With
            wss.Add twss
            
        Next j

        Set pc = dst.Range(s.CurRePo)
        drawWS wss, dst, pc
        s.CurRePo = pc.Offset(wss.Count + 17, 0).Address
    Next
End Sub

Sub 绘制选定多列的风切变图()
    Dim st As Object: Set st = ActiveSheet
    Dim ar: ar = Split(Selection.Address, ",")
    
    Dim wss As New Collection

    Dim cArr As New Collection
    
    For i = 0 To UBound(ar)
        Dim tr As Object: Set tr = st.Range(ar(i))

        If tr.Rows.Count >= tr.Columns.Count Then
            For j = 1 To tr.Columns.Count
                cArr.Add tr.Columns(j)
            Next j
        ElseIf tr.Columns.Count > tr.Rows.Count Then
            For j = 1 To tr.Rows.Count
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
    
    For i = 1 To wss.Count
        dr.Offset(i, 0).Value = wss(i).height
        dr.Offset(i, 1).Value = wss(i).avg
    Next
    
    Dim maxX: maxX = dst.UsedRange.Rows.Count

    Dim rangeH: Set rangeH = dst.Range(dr.Offset(1, 0).Address + ":" + dr.Offset(wss.Count, 0).Address)
    Dim rangeA: Set rangeA = dst.Range(dr.Offset(1, 1).Address + ":" + dr.Offset(wss.Count, 1).Address)

    Dim rangeSort: Set rangeSort = dst.Range(dr.Offset(0, 0).Address + ":" + dr.Offset(wss.Count, 1).Address)

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
    
    For i = 1 To wss.Count
        dr.Offset(0, 1 + i).Value = dr.Offset(i, 0).Value
        
        dr.Offset(i + wss.Count, 0).Value = log(dr.Offset(i, 0).Value)
        dr.Offset(i + wss.Count, 1).Value = log(dr.Offset(i, 1).Value)
        
        For j = 1 To wss.Count
            If dr.Offset(i, 0).Value > dr.Offset(j, 0).Value Then
                dr.Offset(i, j + 1).Value = windshear(dr.Offset(i, 1).Value, dr.Offset(j, 1).Value, dr.Offset(i, 0).Value, dr.Offset(j, 0).Value)
            End If
        Next j
    Next i
    
    Dim range1 As Object: Set range1 = dst.Range(dr.Offset(1 + wss.Count, 1).Address + ":" + dr.Offset(2 * wss.Count, 1).Address)
    Dim range2 As Object: Set range2 = dst.Range(dr.Offset(1 + wss.Count, 0).Address + ":" + dr.Offset(2 * wss.Count, 0).Address)
    Rsq = Application.WorksheetFunction.Rsq(range1, range2)
    Slope = Application.WorksheetFunction.Slope(range1, range2)
    Intercept = Application.WorksheetFunction.Intercept(range1, range2)
    
    Dim a: a = Exp(Intercept)
    Dim b: b = Slope
    
    For i = 1 To wss.Count
        dr.Offset(i, 2 + wss.Count).Value = a * dr.Offset(i, 0).Value ^ b
    Next i
    
    Dim rangeT: Set rangeT = dst.Range(dr.Offset(1, 2 + wss.Count).Address + ":" + dr.Offset(wss.Count, 2 + wss.Count).Address)


    dst.Range(rangeH.Address + "," + rangeA.Address).Select

    Dim myChart As Object: Set myChart = dst.Shapes.AddChart.Chart

    With myChart
        .ChartType = xlXYScatterSmoothNoMarkers
        
        .SetSourceData Source:=Range(dst.Name + "!" + rangeH.Address + "," + dst.Name + "!" + rangeA.Address)

        .SeriesCollection.NewSeries
        .SeriesCollection(2).XValues = "=" + dst.Name + "!" + rangeH.Address
        .SeriesCollection(2).Values = "=" + dst.Name + "!" + rangeT.Address
    
        With .Legend
            .Delete
        End With

        .SetElement (msoElementPrimaryValueAxisTitleRotated)
        .Axes(xlValue).TickLabels.NumberFormatLocal = "0.0_ " '"#,##0.0_);[红色](#,##0.0)"
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "风速 (m/s)"
        .Axes(xlCategory).HasTitle = True
        With .Axes(xlCategory).AxisTitle
            .Format.TextFrame2.TextRange.Characters.Text = "高度 (m)"
        End With
        
        With .SeriesCollection(1)
            .MarkerStyle = -4105
            .Format.Line.Visible = msoFalse
        End With
        
        .Axes(xlCategory).MinimumScale = dr.Offset(1, 0).Value - 5
        
        .SetElement (msoElementChartTitleAboveChart)
        With .ChartTitle
            .Text = "风切变"

            With .Characters.Font
                .Size = 14
            End With
        End With
        
    End With
    
    With myChart.Parent
         .height = 200  ' resize
         .Width = 550   ' resize
         .Top = 0       ' reposition
         .Left = 0      ' reposition
    End With
    
    Dim tb As String: tb = "y = " & Format(a, "0.00") & "x" & Format(b, "0.00")
    Dim tbs As Integer: tbs = InStr(1, tb, "x", 1) + 1
    Dim tbo As Integer: tbo = Len(tb) - tbs + 1


    myChart.Shapes.AddTextbox(msoTextOrientationHorizontal, 417, 91.8, 103.8, _
        56.4).Select
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = tb & Chr(13) & _
                                                "R2 = " & Format(Rsq, "0.00")
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(tbs, tbo).Font
        .BaselineOffset = 0.3
    End With
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(Len(tb) + 3, 1).Font
        .BaselineOffset = 0.3
    End With

    myChart.Parent.Cut
    dst.Select
    dr.Offset(1 + wss.Count, 0).Select
    dst.Pictures.Paste.Select
    
    dst.Range(dr.Offset(1, 1).Address + ":" + dr.Offset(wss.Count, wss.Count + 2).Address).NumberFormatLocal = "0.00"

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
            MsgBox "选定数据不符 " & Selection.Rows.Count & " " & Selection.Columns.Count
            Exit Sub
        End If
        
    ElseIf UBound(ar) = 1 Then
        Dim ta1: ta1 = largethan(st.Range(ar(0)), 1)
        Dim ta2: ta2 = largethan(st.Range(ar(1)), 1)
        Dim r1: Set r1 = st.Range(ar(0))
        Dim r2: Set r2 = st.Range(ar(1))
        
        If ta1 = 1 And ta2 = 1 And r1.Rows.Count = r2.Rows.Count Then
        ElseIf ta1 = 2 And ta2 = 2 And r1.Columns.Count = r2.Columns.Count Then
        Else
            MsgBox "选定数据不符 " & Selection.Rows.Count & " " & Selection.Columns.Count
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
    If o.Rows.Count > i And o.Columns.Count = i Then
        largethan = 1
    ElseIf o.Columns.Count > i And o.Rows.Count = i Then
        largethan = 2
    Else
        largethan = 0
    End If
End Function
