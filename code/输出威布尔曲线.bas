Attribute VB_Name = "输出威布尔曲线"

Function 绘制威布尔曲线(s As Object, rst As Object, dst As Object)
    ' 代表年的不同高度风频曲线及威布尔参数
    s.Pc.Value = "代表年的不同高度风频曲线及威布尔参数"
    s.Pc = s.Pc.Offset(1, 0)
    
    ' 增加数据透视表
    Dim t As Object: Set t = newSheet("tcalweibull")
    Dim pt As Object: Set pt = newPT(t, s.dataRange, "pt")
    
    Dim wvs As Object: Set wvs = s.Sensors("wv")
    Dim a: a = wvs.Items
    For j = 0 To wvs.count - 1
        Dim ss As Object: Set ss = a(j)
        

        showWeibull rst:=rst, dst:=dst, s:=s, t:=t, pt:=pt, ss:=ss
    Next j

    deleteSheet t
End Function


Private Function showWeibull(rst As Object, dst As Object, s As Object, t As Object, pt As Object, ss As Object)
    Dim maxX1 As Integer: maxX1 = rst.UsedRange.Rows.count
    Dim range1 As Object: Set range1 = rst.Range(rst.Cells(2, ss.avg).Address + ":" + rst.Cells(maxX1, ss.avg).Address)

    s.Pc.Value = "CH" + ss.channel + " " + CStr(ss.height) + "m 代表年威布尔曲线图"
    Dim po As Object: Set po = s.Pc.Offset(1, 0)

    Dim sr As Double: sr = Sqr(Application.WorksheetFunction.DevSq(range1) / (maxX1 - 1))
    Dim avg As Double: avg = Application.WorksheetFunction.Average(range1)

    Dim k As Double: k = (avg / sr) ^ 1.086
    Dim c As Double: c = avg / gamma(1 + 1 / k)

    pt.ClearTable
    
    With pt.PivotFields("CH" + ss.channel + "Wb")
        .Orientation = xlColumnField
        .Position = 1
    End With


    Dim ssn As String: ssn = "风速频率"

    pt.AddDataField pt.PivotFields("CH" + ss.channel + "Avg"), ssn, xlCount
    
    With pt.PivotFields(ssn)
        .Calculation = xlPercentOfTotal
        .NumberFormat = "0.00%"
    End With
    
    With pt
        .ColumnGrand = False
        .RowGrand = False
    End With
    

    Dim maxX As Integer: maxX = t.UsedRange.Rows.count
    Dim maxY As Integer: maxY = t.UsedRange.Columns.count
    rangeCopy t.Range("A2:" + t.Cells(maxX, maxY).Address), po

    po.Value = "风速 (m/s)"
    po.Offset(1, 0).Value = "风速频率 (%)"
    
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
    
    Dim rangeX As String: rangeX = dst.Name + "!" + po.Offset(2, 1).Address + ":" + po.Offset(2, tdmax).Address
    Dim cRangeY As New Collection, cRangeT As New Collection
    For i = 3 To 4
        cRangeY.Add po.Offset(i, 1).Address + ":" + po.Offset(i, tdmax).Address
        cRangeT.Add dst.Name + "!" + po.Offset(i, 0).Address
    Next i
    
    Dim myChart As Object
    Set myChart = drawChart(rangeX:=rangeX, cRangeY:=cRangeY, cRangeT:=cRangeT, rst:=dst, dst:=dst, _
            dpo:=po.Offset(2, 0), axisTitleX:="风速 (m/s)", axisTitleY:="频率 (%)", axisFormatX:="0", _
            tickSpacingX:=100)
    
    myChart.SeriesCollection(1).ChartType = xlArea

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

    'dst.Range(po.Offset(2, 0).Address + ":" + po.Offset(4, tdmax).Address).Clear
    
    Dim range2 As Object: Set range2 = dst.Range(po.Offset(1, 1), po.Offset(1, maxY - 1))
    rangeFV dr:=range2, v:=100, m:=xlMultiply
    range2.NumberFormatLocal = "0.00"

    s.Pc = po.Offset(2 + 16, 0)
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

