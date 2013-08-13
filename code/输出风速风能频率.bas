Attribute VB_Name = "输出风速风能频率"

Sub 计算风速风能频率()
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
            s.DataRange, Version:=xlPivotTableVersion14). _
            CreatePivotTable TableDestination:=oTemp.Name + "!R1C1", TableName:="pt", _
            DefaultVersion:=xlPivotTableVersion14
        Dim pt As Object: Set pt = oTemp.PivotTables("pt")

        ' 不同高度风速和风能频率分布
        Dim pc As Object: Set pc = dst.Range(s.CurRePo)
        pc.Value = "不同高度风速和风能频率分布"
        s.CurRePo = pc.Offset(1, 0).Address
        
        Dim wvs As Object: Set wvs = s.Sensors("wv")
        Dim a: a = wvs.Items
        For j = 0 To wvs.Count - 1
            Dim ss As Object: Set ss = a(j)
            
            Set pc = dst.Range(s.CurRePo)
            showWfv rst:=rst, dst:=dst, s:=s, po:=pc, pt:=pt, ss:=ss
        Next

        ' 清除数据透视表，删除增加的数据列
        oTemp.Range(pt.TableRange2.Address).Delete Shift:=xlUp
    Next
End Sub

Private Function showWfv(rst As Object, dst As Object, s As Object, po As Object, pt As Object, ss As Object)
    ' 代表年不同高度风速和风能频率分布

    pt.ClearTable

    Dim wfvn As String: wfvn = "CH" + ss.channel + "Wfv"
    With pt.PivotFields(wfvn)
        .Orientation = xlRowField
        .Position = 1
    End With

    Dim wvn As String: wvn = ss.channel + " " + CStr(ss.height) + "m " + "风速频率"
    Dim wpn As String: wpn = ss.channel + " " + CStr(ss.height) + "m " + "风功率频率"
    pt.AddDataField pt.PivotFields("CH" + ss.channel + "Avg"), _
        wvn, xlCount
    pt.AddDataField pt.PivotFields("CH" + ss.channel + "WP"), _
        wpn, xlSum
    
    With pt.DataPivotField
        .Orientation = xlColumnField
        .Position = 1
    End With
    
    With pt.PivotFields(wvn)
        .Calculation = xlPercentOfTotal
        .NumberFormat = "0.00%"
    End With

    With pt.PivotFields(wpn)
        .Calculation = xlPercentOfTotal
        .NumberFormat = "0.00%"
    End With

    With pt
        .ColumnGrand = False
        .RowGrand = False
    End With
    
    Dim maxX As Integer: maxX = oTemp.UsedRange.Rows.Count
    Dim maxY As Integer: maxY = oTemp.UsedRange.Columns.Count
    oTemp.UsedRange.Copy
    po.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    If po.Offset(1, 0).Value = 0.5 Then
        po.Offset(1, 0).Value = "≤0.5"
    End If

    '乘100显示
    Dim t100 As Object: Set t100 = po.Offset(0, maxY)
    t100.Value = 100
    t100.Copy
    With dst.Range(po.Offset(1, 1).Address + ":" + po.Offset(maxX - 1, maxY - 1).Address)
        .PasteSpecial Paste:=xlPasteAll, Operation:=xlMultiply, _
                      SkipBlanks:=False, Transpose:=False
        .NumberFormatLocal = "0.00_ "
    End With
    t100.Clear

    po.Value = "风速区间(m/s)"
    
    Dim range2 As Object: Set range2 = dst.Range(po.Offset(0, 1).Address + ":" + po.Offset(maxX - 1, maxY - 1).Address)
    Dim range3 As String: range3 = po.Offset(1, 0).Address + ":" + po.Offset(maxX - 1, 0).Address
    Dim myChart As Object: Set myChart = dst.Shapes.AddChart.Chart
    With myChart
        .ChartType = xlColumnClustered
        .SetSourceData Source:=range2
        
        With .Legend
            .Position = xlTop
        End With
        
        .SetElement (msoElementPrimaryValueAxisTitleRotated)
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "频率 (%)"
        .Axes(xlCategory).HasTitle = True
        With .Axes(xlCategory).AxisTitle
            .Format.TextFrame2.TextRange.Characters.Text = "风速 (m/s)"
        End With
        
        .SeriesCollection(1).XValues = "=" + dst.Name + "!" + range3
        .Axes(xlValue).TickLabels.NumberFormatLocal = "0_ "
    End With
    With myChart.Parent
         .height = 200  ' resize
         .Width = 550   ' resize
         .Top = 0       ' reposition
         .Left = 0      ' reposition
    End With

    myChart.Parent.Cut
    dst.Select
    po.Offset(maxX, 0).Select
    dst.Pictures.Paste.Select
    
    Dim range1 As Object: Set range1 = dst.Range(po.Offset(1, 0).Address + ":" + po.Offset(maxX - 1, 0).Address)
    range1.NumberFormatLocal = "0.0_ "
    
    pt.ClearTable
    
    s.CurRePo = po.Offset(maxX + 16, 0).Address

End Function

