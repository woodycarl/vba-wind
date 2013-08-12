Attribute VB_Name = "输出平均风速风能"

Sub 计算平均风速风能()
    系统初始化
    
    For Each k In Stations
        Dim s As Object: Set s = Stations(k)
        
        
        If s.CurRePo = "A1" Then
            initCal s
            
            s.CurRePo = "A3"
            
        End If
        
        Dim po As Object: Set po = Sheets("result" + s.id).Range(s.CurRePo)


        showResult s
    Next
End Sub

Function showAvg(rst As Object, dst As Object, s As Object, po As Object, pt As Object, unit As String, cat As String)
    ' 代表年不同高度月平均

    pt.ClearTable
    
    With pt.PivotFields("Month")
        .Orientation = xlColumnField
        .Position = 1
    End With

    Dim wvs As Object: Set wvs = s.Sensors("wv")
    Dim a: a = wvs.Items
    For j = 0 To wvs.Count - 1
        Dim ss As Object: Set ss = a(j)
        Dim ssn As String: ssn = "CH" + ss.channel + cat

        pt.AddDataField pt.PivotFields(ssn), ss.channel + " " + CStr(ss.height) + "m", xlAverage
    Next
    
    With pt.DataPivotField
        .Orientation = xlRowField
        .Position = 1
    End With

    Dim offsetX As Integer: offsetX = oTemp.UsedRange.Rows.Count - 2
    Dim offsetY As Integer: offsetY = oTemp.UsedRange.Columns.Count - 1
    oTemp.Range("A2:" + oTemp.Range("A2").Offset(offsetX, offsetY).Address).Copy
    po.Offset(1, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Dim range3 As Object: Set range3 = dst.Range(po.Offset(2, 2).Address + ":" + po.Offset(offsetX + 2, offsetY + 2).Address)
    range3.NumberFormatLocal = "0.00_ "

    With dst.Range(po.Offset(1, 0).Address + ":" + po.Offset(1, 1).Address)
        .HorizontalAlignment = xlCenter
        .Merge
        .Value = "时间 (月)"
    End With
    
    With dst.Range(po.Offset(2, 0).Address + ":" + po.Offset(wvs.Count + 1, 0).Address)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Merge
    End With
    po.Offset(2, 0).Value = unit
    
    po.Offset(1, offsetY + 1).Value = "平均"

    Dim range4 As Object: Set range4 = dst.Range(po.Offset(2, 1).Address + ":" + po.Offset(offsetX + 1, offsetY).Address)
    Dim range5 As String: range5 = po.Offset(1, 2).Address + ":" + po.Offset(1, offsetY).Address
    Dim myChart As Object: Set myChart = dst.Shapes.AddChart.Chart
    With myChart
        .ChartType = xlLine
        .SetSourceData Source:=range4
        
        With .Legend
            .Position = xlTop
        End With
        
        .SetElement (msoElementPrimaryValueAxisTitleRotated)
        .Axes(xlValue, xlPrimary).AxisTitle.Text = unit
        .Axes(xlCategory).HasTitle = True
        With .Axes(xlCategory).AxisTitle
            .Format.TextFrame2.TextRange.Characters.Text = "月份"
        End With
        
        .SeriesCollection(1).XValues = "=" + dst.Name + "!" + range5
    End With
    With myChart.Parent
         .height = 200  ' resize
         .Width = 550   ' resize
         .Top = 0       ' reposition
         .Left = 0      ' reposition
    End With

    myChart.Parent.Cut
    dst.Select
    po.Offset(2 + wvs.Count, 0).Select
    dst.Pictures.Paste.Select
    
    pt.ClearTable
    
End Function

Function showAvgH(rst As Object, dst As Object, s As Object, po As Object, pt As Object, unit As String, cat As String)
    ' 代表年不同高度小时平均

    pt.ClearTable

    With pt.PivotFields("Hour")
        .Orientation = xlRowField
        .Position = 1
    End With

    Dim wvs As Object: Set wvs = s.Sensors("wv")
    Dim a: a = wvs.Items
    For j = 0 To wvs.Count - 1
        Dim ss As Object: Set ss = a(j)
        Dim ssn As String: ssn = "CH" + ss.channel + cat

        pt.AddDataField pt.PivotFields(ssn), ss.channel + " " + CStr(ss.height) + "m", xlAverage
    Next
    
    With pt.DataPivotField
        .Orientation = xlColumnField
        .Position = 1
    End With

    Dim maxX As Integer: maxX = oTemp.UsedRange.Rows.Count
    Dim maxY As Integer: maxY = oTemp.UsedRange.Columns.Count
    oTemp.UsedRange.Copy
    po.Offset(2, 0).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Dim range3 As Object: Set range3 = dst.Range(po.Offset(3, 1).Address + ":" + po.Offset(maxX + 1, maxY - 1).Address)
    range3.NumberFormatLocal = "0.00_ "

    With dst.Range(po.Offset(1, 0).Address + ":" + po.Offset(2, 0).Address)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Merge
        .Value = "时间 (小时)"
    End With
    
    With dst.Range(po.Offset(1, 1).Address + ":" + po.Offset(1, wvs.Count + 1).Address)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Merge
        .Value = unit
    End With

    po.Offset(maxX + 1, 0).Value = "平均"

    Dim range4 As Object: Set range4 = dst.Range(po.Offset(2, 1).Address + ":" + po.Offset(maxX, maxY - 1).Address)
    Dim range5 As String: range5 = po.Offset(3, 0).Address + ":" + po.Offset(maxX, 0).Address
    Dim myChart As Object: Set myChart = dst.Shapes.AddChart.Chart
    With myChart
        .ChartType = xlLine
        .SetSourceData Source:=range4
        
        With .Legend
            .Position = xlTop
        End With
        
        .SetElement (msoElementPrimaryValueAxisTitleRotated)
        .Axes(xlValue, xlPrimary).AxisTitle.Text = unit
        .Axes(xlCategory).HasTitle = True
        With .Axes(xlCategory).AxisTitle
            .Format.TextFrame2.TextRange.Characters.Text = "小时"
        End With
        
        .SeriesCollection(1).XValues = "=" + dst.Name + "!" + range5
    End With
    With myChart.Parent
         .height = 200  ' resize
         .Width = 550   ' resize
         .Top = 0       ' reposition
         .Left = 0      ' reposition
    End With

    myChart.Parent.Cut
    dst.Select
    po.Offset(28, 0).Select
    dst.Pictures.Paste.Select
    
    pt.ClearTable
    
End Function

