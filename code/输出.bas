Attribute VB_Name = "输出"

Sub 生成报表()
    系统初始化
    
    For Each k In Stations
        Dim s As Object: Set s = Stations(k)
        'initCal s
        showResult s
    Next
    
End Sub

Private Function showResult(s As Object)
    Dim rst As Object: Set rst = Sheets(s.Sheet1h)
    
    Dim dst As Object: Set dst = Sheets.Add(after:=Sheets(Sheets.Count))
    dst.Name = "result" + s.id
    
    dst.Range("A1").Value = "数据日期" + Format(s.StartTime, "yyyy年mm月dd日") + "～" + Format(s.EndTime, "yyyy年mm月dd日")


    ' 1
    dst.Range("A2").Value = "1、代表年不同高度月平均风速"
    
    With dst.Range("A3:B3")
        .HorizontalAlignment = xlCenter
        .Merge
        .Value = "时间 (月)"
    End With
    
    Dim wvs As Object: Set wvs = s.Sensors("wv")
    Dim a: a = wvs.Items

    rst.Columns("A:A").NumberFormatLocal = "m"
    rst.Range("A1").AutoFilter
    
    Dim i As Integer
    Dim indexy As Integer: indexy = 1
    For i = 1 To 12
        rst.UsedRange.AutoFilter Field:=1, Criteria1:="=" + CStr(i), _
            Operator:=xlAnd
        
        If rst.Range("A1").CurrentRegion.Rows.Count <= 2 Then
            GoTo notEnouthData
        End If
        
        dst.Cells(3, indexy + 2).Value = i
        
        oTemp.UsedRange.Clear
        rst.Range("A1").CurrentRegion.Copy
        oTemp.Select
        Range("A1").Select
        oTemp.Paste
        Range("A1").Select
    
        For j = 0 To wvs.Count - 1
            Dim ss As Object: Set ss = a(j)

            If dst.Cells(4 + j, 2).Value = "" Then
                dst.Cells(4 + j, 2).Value = CStr(ss.Height) + "m"  '"CH" + ss.channel + " " +
            End If
            
            Dim rangei As Object: Set rangei = oTemp.Range(arrCol(CInt(ss.channel)))
            Dim avg As Double: avg = Application.WorksheetFunction.Average(rangei)
            
            dst.Cells(4 + j, indexy + 2).Value = avg
        Next
        
notEnouthData:
        indexy = indexy + 1
    Next i
    
    dst.Cells(3, indexy + 2).Value = "平均"
    For j = 0 To wvs.Count - 1
        Dim rangej As Object: Set rangej = dst.Range(dst.Cells(4 + j, 3).Address + ":" + dst.Cells(4 + j, indexy + 1).Address)
        dst.Cells(4 + j, indexy + 2).Value = Application.WorksheetFunction.Average(rangej)
    Next
    rst.Range("A1").AutoFilter
    
    Dim range3 As Object: Set range3 = dst.Range(dst.Cells(4, 3).Address + ":" + dst.Cells(3 + wvs.Count, indexy + 2).Address)
    range3.NumberFormatLocal = "0.00_ "

    With dst.Range("A4:A" + CStr(j + 3))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Merge
    End With
    dst.Range("A4").Value = "风速 (m/s)"

    Dim range4 As Object: Set range4 = dst.Range(dst.Cells(4, 2).Address + ":" + dst.Cells(3 + wvs.Count, indexy + 1).Address)
    Dim range5 As String: range5 = dst.Cells(3, 3).Address + ":" + dst.Cells(3, indexy + 1).Address
    Dim wvavgChart As Object: Set wvavgChart = dst.Shapes.AddChart.Chart
    With wvavgChart
        .ChartType = xlLine
        .SetSourceData Source:=range4
        
        With .Legend
            .Position = xlTop
        End With
        
        .SetElement (msoElementPrimaryValueAxisTitleRotated)
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "风速 (m/s)"
        .Axes(xlCategory).HasTitle = True
        With .Axes(xlCategory).AxisTitle
            .Format.TextFrame2.TextRange.Characters.Text = "月份"
        End With
        
        .SeriesCollection(1).XValues = "=" + dst.Name + "!" + range5
    End With
    With wvavgChart.Parent
         .Height = 200  ' resize
         .Width = 550   ' resize
         .Top = 0       ' reposition
         .Left = 0      ' reposition
    End With

    wvavgChart.Parent.Cut
    dst.Select
    dst.Cells(5 + j, 1).Select
    dst.Pictures.Paste.Select
    
    
    '2
    dst.Cells(20 + j, 1).Value = "2、代表年不同高度月平均风功率密度"
    
    
    
End Function

Sub testCal()
    系统初始化
    
    For Each k In Stations
        Dim s As Object: Set s = Stations(k)
        initCal s

    Next
End Sub

Private Function initCal(s As Object)
    calAirDensity s
    
End Function

Private Function calAirDensity(s As Object)
    Dim ts As Object: Set ts = s.Sensors("t")
    Dim ps As Object: Set ps = s.Sensors("p")
    
    Dim st As Object: Set st = Sheets(s.Sheet1h)
    
    Dim p, t
    If ts.Count > 0 Then
        Dim ta: ta = ts.Items
        Dim sst As Object: Set sst = ta(0)
        Dim ranget As Object: Set ranget = st.Range(arrCol(sst.channel))
        t = Application.WorksheetFunction.Average(ranget)
        
        If ps.Count > 0 Then
            Dim pa: pa = ps.Items
            Dim ssp As Object: Set ssp = pa(0)
            Dim rangep As Object: Set rangep = st.Range(arrCol(ssp.channel))
            p = Application.WorksheetFunction.Average(rangep)
            
            s.AirDensity = p * 1000 / (287 * (t + 273))
            
        Else
            s.AirDensity = (353.05 / t) * Exp((-0.034) * (Height / (t + 273)))
        End If
    Else
        s.AirDensity = oConfig.AirDensity
    End If
End Function

Private Function calTurbs(s As Object)

End Function
