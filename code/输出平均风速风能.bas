Attribute VB_Name = "输出平均风速风能"
Private unitDic As Scripting.Dictionary

Sub 计算平均风速风能()
    系统初始化
    init
    
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

        ' 1、代表年不同高度月平均风速
        Dim pc As Object: Set pc = dst.Range(s.CurRePo)
        pc.Value = "代表年不同高度月平均风速"
        showAvg rst:=rst, dst:=dst, s:=s, po:=pc, pt:=pt, cat:="Avg"
    
        ' 2、代表年不同高度月平均风功率密度
        Set pc = dst.Range(s.CurRePo)
        pc.Value = "代表年不同高度月平均风功率密度"
        showAvg rst:=rst, dst:=dst, s:=s, po:=pc, pt:=pt, cat:="WP"
    
        ' 3、代表年不同高度小时平均风速
        Set pc = dst.Range(s.CurRePo)
        pc.Value = "代表年不同高度小时平均风速"
        showAvgH rst:=rst, dst:=dst, s:=s, po:=pc, pt:=pt, cat:="Avg"
    
        ' 4、代表年不同高度小时平均风功率密度
        Set pc = dst.Range(s.CurRePo)
        pc.Value = "代表年不同高度小时平均风功率密度"
        showAvgH rst:=rst, dst:=dst, s:=s, po:=pc, pt:=pt, cat:="WP"
        
        ' 清除数据透视表，删除增加的数据列
        oTemp.Range(pt.TableRange2.Address).Delete Shift:=xlUp
        'rst.Range(s.AddedData).Clear
    Next
End Sub

Private Function init()
    Set unitDic = CreateObject("Scripting.Dictionary")
    unitDic.Add "Avg", "风速 (m/s)"
    unitDic.Add "WP", "风功率密度 (W/m2)"
End Function

Private Function showAvg(rst As Object, dst As Object, s As Object, po As Object, pt As Object, cat As String)
    ' 代表年不同高度月平均
    Dim unit As String: unit = unitDic(cat)

    pt.ClearTable
    
    With pt.PivotFields("Month")
        .Orientation = xlColumnField
        .Position = 1
    End With

    Dim wvs As Object: Set wvs = s.Sensors("wv")
    Dim a: a = wvs.Items
    For j = 0 To wvs.count - 1
        Dim ss As Object: Set ss = a(j)
        Dim ssn As String: ssn = "CH" + ss.channel + cat

        pt.AddDataField pt.PivotFields(ssn), ss.channel + " " + CStr(ss.height) + "m", xlAverage
    Next
    
    With pt.DataPivotField
        .Orientation = xlRowField
        .Position = 1
    End With

    Dim maxX As Integer: maxX = oTemp.UsedRange.Rows.count
    Dim maxY As Integer: maxY = oTemp.UsedRange.Columns.count
    oTemp.Range("A2:" + oTemp.Cells(maxX, maxY).Address).Copy
    po.Offset(1, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    With dst.Range(po.Offset(2, 2).Address + ":" + po.Offset(maxX, maxY + 1).Address)
        .NumberFormatLocal = "0.00_ "
    End With

    With dst.Range(po.Offset(1, 0).Address + ":" + po.Offset(1, 1).Address)
        .HorizontalAlignment = xlCenter
        .Merge
        .Value = "时间 (月)"
    End With
    
    With dst.Range(po.Offset(2, 0).Address + ":" + po.Offset(wvs.count + 1, 0).Address)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Merge
        .Value = unit
    End With

    po.Offset(1, maxY).Value = "平均"
    
    Dim range2 As Object: Set range2 = dst.Range(po.Offset(2, 1).Address + ":" + po.Offset(maxX - 1, maxY - 1).Address)
    Dim range3 As String: range3 = po.Offset(1, 2).Address + ":" + po.Offset(1, maxY - 1).Address
    Dim myChart As Object: Set myChart = dst.Shapes.AddChart.Chart
    With myChart
        .ChartType = xlLine
        .SetSourceData Source:=range2
        
        With .Legend
            .Position = xlTop
        End With
        
        .SetElement (msoElementPrimaryValueAxisTitleRotated)
        .Axes(xlValue).TickLabels.NumberFormatLocal = "0_ "
        .Axes(xlValue, xlPrimary).AxisTitle.Text = unit
        .Axes(xlCategory).HasTitle = True
        With .Axes(xlCategory).AxisTitle
            .Format.TextFrame2.TextRange.Characters.Text = "月份"
        End With
        
        .SeriesCollection(1).XValues = "=" + dst.Name + "!" + range3
    End With
    With myChart.Parent
         .height = 200  ' resize
         .Width = 550   ' resize
         .Top = 0       ' reposition
         .Left = 0      ' reposition
    End With

    myChart.Parent.Cut
    dst.Select
    po.Offset(2 + wvs.count, 0).Select
    dst.Pictures.Paste.Select
    
    pt.ClearTable
    
    s.CurRePo = po.Offset(2 + wvs.count + 16, 0).Address
End Function

Private Function showAvgH(rst As Object, dst As Object, s As Object, po As Object, pt As Object, cat As String)
    ' 代表年不同高度小时平均
    Dim unit As String: unit = unitDic(cat)
    
    pt.ClearTable

    With pt.PivotFields("Hour")
        .Orientation = xlRowField
        .Position = 1
    End With

    Dim wvs As Object: Set wvs = s.Sensors("wv")
    Dim a: a = wvs.Items
    For j = 0 To wvs.count - 1
        Dim ss As Object: Set ss = a(j)
        Dim ssn As String: ssn = "CH" + ss.channel + cat

        pt.AddDataField pt.PivotFields(ssn), ss.channel + " " + CStr(ss.height) + "m", xlAverage
    Next
    
    With pt.DataPivotField
        .Orientation = xlColumnField
        .Position = 1
    End With

    Dim maxX As Integer: maxX = oTemp.UsedRange.Rows.count
    Dim maxY As Integer: maxY = oTemp.UsedRange.Columns.count
    oTemp.UsedRange.Copy
    po.Offset(2, 0).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    With dst.Range(po.Offset(3, 1).Address + ":" + po.Offset(maxX + 1, maxY - 1).Address)
        .NumberFormatLocal = "0.00_ "
    End With
    
    With dst.Range(po.Offset(1, 0).Address + ":" + po.Offset(2, 0).Address)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Merge
        .Value = "时间 (小时)"
    End With
    
    With dst.Range(po.Offset(1, 1).Address + ":" + po.Offset(1, wvs.count).Address)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Merge
        .Value = unit
    End With

    po.Offset(maxX + 1, 0).Value = "平均"

    Dim range2 As Object: Set range2 = dst.Range(po.Offset(2, 1).Address + ":" + po.Offset(maxX, maxY - 1).Address)
    Dim range3 As String: range3 = po.Offset(3, 0).Address + ":" + po.Offset(maxX, 0).Address
    Dim myChart As Object: Set myChart = dst.Shapes.AddChart.Chart
    With myChart
        .ChartType = xlLine
        .SetSourceData Source:=range2
        
        With .Legend
            .Position = xlTop
        End With
        
        .SetElement (msoElementPrimaryValueAxisTitleRotated)
        .Axes(xlValue).TickLabels.NumberFormatLocal = "0_ "
        .Axes(xlValue, xlPrimary).AxisTitle.Text = unit
        .Axes(xlCategory).HasTitle = True
        With .Axes(xlCategory).AxisTitle
            .Format.TextFrame2.TextRange.Characters.Text = "小时"
        End With
        
        .SeriesCollection(1).XValues = "=" + dst.Name + "!" + range3
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
    
    s.CurRePo = po.Offset(28 + 16, 0).Address
End Function

