Attribute VB_Name = "输出威布尔曲线"

Sub 绘制威布尔曲线()
    系统初始化
    
    For Each k In Stations
        Dim s As Object: Set s = Stations(k)
        
        If s.CurRePo = "A1" Then
            initCalResult s
        End If
        
        Dim rst As Object: Set rst = Sheets(s.Sheet1h)
        Dim dst As Object: Set dst = Sheets(s.Result)

        oTemp.UsedRange.Clear
        
        ' 增加数据透视表
        oWB.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
            s.dataRange, Version:=xlPivotTableVersion14). _
            CreatePivotTable TableDestination:=oTemp.Name + "!R1C1", TableName:="pt", _
            DefaultVersion:=xlPivotTableVersion14
        Dim pt As Object: Set pt = oTemp.PivotTables("pt")

        ' 代表年的不同高度风频曲线及威布尔参数
        Dim pc As Object: Set pc = dst.Range(s.CurRePo)
        pc.Value = "代表年的不同高度风频曲线及威布尔参数"
        s.CurRePo = pc.Offset(1, 0).Address
        
        Dim wvs As Object: Set wvs = s.Sensors("wv")
        Dim a: a = wvs.Items
        For j = 0 To wvs.count - 1
            Dim ss As Object: Set ss = a(j)
            
            Set pc = dst.Range(s.CurRePo)
            showWeibull rst:=rst, dst:=dst, s:=s, po:=pc, pt:=pt, ss:=ss
        Next j

        ' 清除数据透视表，删除增加的数据列
        oTemp.Range(pt.TableRange2.Address).Delete Shift:=xlUp
        
    Next
End Sub


Private Function showWeibull(rst As Object, dst As Object, s As Object, po As Object, pt As Object, ss As Object)
    Dim maxX1 As Integer: maxX1 = rst.UsedRange.Rows.count
    Dim range1 As Object: Set range1 = rst.Range(rst.Cells(2, ss.avg).Address + ":" + rst.Cells(maxX1, ss.avg).Address)

    Dim sr As Double
    sr = Sqr(Application.WorksheetFunction.DevSq(range1) / (maxX1 - 1))
    Dim avg As Double
    avg = Application.WorksheetFunction.Average(range1)

    Dim k As Double, c As Double
    k = (avg / sr) ^ 1.086
    c = avg / gamma(1 + 1 / k)

    
    pt.ClearTable
    
    With pt.PivotFields("CH" + ss.channel + "Wb")
        .Orientation = xlColumnField
        .Position = 1
    End With


    Dim ssn As String: ssn = ss.channel + " " + CStr(ss.height) + "m"

    pt.AddDataField pt.PivotFields("CH" + ss.channel + "Avg"), ssn, xlCount
    
    With pt.PivotFields(ssn)
        .Calculation = xlPercentOfTotal
        .NumberFormat = "0.00%"
    End With
    
    With pt
        .ColumnGrand = False
        .RowGrand = False
    End With
    

    Dim maxX As Integer: maxX = oTemp.UsedRange.Rows.count
    Dim maxY As Integer: maxY = oTemp.UsedRange.Columns.count
    oTemp.Range("A2:" + oTemp.Cells(maxX, maxY).Address).Copy
    po.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    po.Value = "风速 (m/s)"
    po.Offset(1, 0).Value = ssn + " 频率 (%)"
    
    po.Offset(2, 0).Value = "风速"
    po.Offset(3, 0).Value = "风速频率"
    po.Offset(4, 0).Value = "威布尔曲线"
    
    po.Offset(2, 1).Value = 0
    po.Offset(3, 1).Value = 0 'po.Offset(1, 1).Value
    po.Offset(4, 1).Value = weibull(k, c, 0)
    
    For i = 1 To maxY - 1
        For j = 1 To 100
            Dim v As Double: v = 1 / 100 * j + i - 1
            Dim x As Integer: x = (i - 1) * 100 + j + 1
            po.Offset(2, x).Value = v
            If j = 100 Then 'j < 3 Or j > 97
                po.Offset(3, x).Value = 0
            Else
                po.Offset(3, x).Value = 100 * po.Offset(1, i).Value
            End If
            
            po.Offset(4, x).Value = 100 * weibull(k, c, v)
            
        Next j
    Next i
    
    Dim tdmax As Integer: tdmax = (maxY - 1) * 100 + 1
    
    Dim range2 As Object: Set range2 = dst.Range(po.Offset(3, 0).Address + ":" + po.Offset(4, tdmax).Address)
    Dim range3 As String: range3 = po.Offset(2, 1).Address + ":" + po.Offset(2, tdmax).Address
    Dim myChart As Object: Set myChart = dst.Shapes.AddChart.Chart
    With myChart
        .ChartType = xlLine
        .SetSourceData Source:=range2
        
        With .Legend
            .Position = xlTop
        End With
        
        .Axes(xlCategory).TickLabelSpacing = 100
        .Axes(xlCategory).TickMarkSpacing = 100
        .SetElement (msoElementPrimaryValueAxisTitleRotated)
        .Axes(xlValue).TickLabels.NumberFormatLocal = "0_ "
        .Axes(xlValue, xlPrimary).AxisTitle.Text = unit
        .Axes(xlCategory).HasTitle = True
        With .Axes(xlCategory).AxisTitle
            .Format.TextFrame2.TextRange.Characters.Text = "风速 (m/s)"
        End With
        
        .SetElement (msoElementPrimaryValueAxisTitleRotated)
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "频率 (%)"
        
        .SeriesCollection(1).XValues = "=" + dst.Name + "!" + range3
    End With

    With myChart.Parent
         .height = 200  ' resize
         .Width = 550   ' resize
         .Top = 0       ' reposition
         .Left = 0      ' reposition
    End With
    
    Dim wpavg As Double: wpavg = Application.WorksheetFunction.Average(rst.Range(wpRange(rst, ss.channel)))

    With myChart.Shapes.AddTextbox(msoTextOrientationHorizontal, 400, 50, 100, 100)
        With .TextFrame2.TextRange
            .Characters.Text = "A: " & Format(c, "0.00") & " m/s" & Chr(13) & _
                               "k: " & Format(k, "0.00") & Chr(13) & _
                               "U: " & Format(avg, "0.00") & " m/s" & Chr(13) & _
                               "P: " & Format(wpavg, "0.00") & " W/m2"
            With .Characters.Font
                .Size = 11
            End With
        End With

    End With
    
    myChart.Parent.Cut
    dst.Select
    po.Offset(2, 0).Select
    dst.Pictures.Paste.Select
    
    dst.Range(po.Offset(2, 0).Address + ":" + po.Offset(4, tdmax).Address).Clear
    '乘100显示 提取为函数
    Dim t100 As Object: Set t100 = po.Offset(0, maxY)
    t100.Value = 100
    t100.Copy
    With dst.Range(po.Offset(1, 1).Address + ":" + po.Offset(1, maxY - 1).Address)
        .PasteSpecial Paste:=xlPasteAll, Operation:=xlMultiply, _
                      SkipBlanks:=False, Transpose:=False
        .NumberFormatLocal = "0.00_ "
    End With
    t100.Clear
    
    s.CurRePo = po.Offset(2 + 16, 0).Address

End Function


Private Function gamma(v As Double) As Double
    gamma = Exp(Application.WorksheetFunction.GammaLn(v))
End Function

Private Function weibull(k As Double, c As Double, v As Double) As Double
    weibull = Application.WorksheetFunction.weibull(v, k, c, False)
End Function

Function wpRange(rst As Object, ch As String) As String
    Dim maxX As Integer: maxX = rst.UsedRange.Rows.count
    Dim maxY As Integer: maxY = rst.UsedRange.Columns.count
    
    For i = 1 To maxY
        If rst.Cells(1, i).Value = "CH" + ch + "WP" Then
            wpRange = rst.Cells(2, i).Address + ":" + rst.Cells(maxX, i).Address
            Exit Function
        End If
    Next
End Function

