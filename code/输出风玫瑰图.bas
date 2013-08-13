Attribute VB_Name = "输出风玫瑰图"

Sub 绘制风玫瑰图()
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

        ' 代表年的全年风向、风能频率分布玫瑰图
        Dim pc As Object: Set pc = dst.Range(s.CurRePo)
        pc.Value = "代表年的全年风向、风能频率分布玫瑰图"
        s.CurRePo = pc.Offset(1, 0).Address
        
        Dim wds As Object: Set wds = s.Sensors("wd")
        Dim ad: ad = wds.Items
        
        Dim wvs As Object: Set wvs = s.Sensors("wv")
        Dim av: av = wvs.Items
        
        For i = 0 To wvs.Count - 1
            Dim ssv As Object: Set ssv = av(i)
            Dim ssd As Object: Set ssd = getSSbyH(wds, ssv.height)
            
            Set pc = dst.Range(s.CurRePo)
            showWindrose rst, dst, s, pc, pt, ssv, ssd
        Next

        ' 清除数据透视表，删除增加的数据列
        oTemp.Range(pt.TableRange2.Address).Delete Shift:=xlUp
    Next
End Sub


Private Function showWindrose(rst As Object, dst As Object, s As Object, po As Object, pt As Object, ssv As Object, ssd As Object)
    ' 代表年的16个方位扇区不同高度出现风向、风能频率
    Dim arrWr: arrWr = Array("N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE", "S", "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW", "N")
    pt.ClearTable
    
    With pt.PivotFields("CH" + ssd.channel + "Wr")
        .Orientation = xlColumnField
        .Position = 1
    End With
    
    Dim ssvn As String: ssvn = ssv.channel + " " + CStr(ssv.height) + "m " + "风向频率"
    Dim sspn As String: sspn = ssv.channel + " " + CStr(ssv.height) + "m " + "风能频率"
    
    pt.AddDataField pt.PivotFields("CH" + ssv.channel + "Avg"), ssvn, xlCount
    pt.AddDataField pt.PivotFields("CH" + ssv.channel + "WP"), sspn, xlSum
    
    With pt.PivotFields(ssvn)
        .Calculation = xlPercentOfTotal
        .NumberFormat = "0.00%"
    End With
    
    With pt.PivotFields(sspn)
        .Calculation = xlPercentOfTotal
        .NumberFormat = "0.00%"
    End With
    
    With pt.DataPivotField
        .Orientation = xlRowField
        .Position = 1
    End With

    With pt
        .ColumnGrand = False
        .RowGrand = False
    End With
    
    Dim maxX As Integer: maxX = oTemp.UsedRange.Rows.Count
    Dim maxY As Integer: maxY = oTemp.UsedRange.Columns.Count
    
    oTemp.Range("A2:" + oTemp.Cells(maxX, maxY).Address).Copy
    
    po.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    po.Value = ""
    For i = 1 To maxY - 1
        po.Offset(0, i).Value = arrWr(po.Offset(0, i).Value - 1)
    Next
    
    '乘100显示
    po.Offset(0, maxY).Value = 100
    po.Offset(0, maxY).Copy
    Dim range1 As Object
    Set range1 = dst.Range(po.Offset(1, 1).Address + ":" + po.Offset(maxX - 2, maxY - 1).Address)
    range1.PasteSpecial Paste:=xlPasteAll, Operation:=xlMultiply, _
        SkipBlanks:=False, Transpose:=False
    range1.NumberFormatLocal = "0.00_ "
    po.Offset(0, maxY).Clear
    
    Dim range2 As Object: Set range2 = dst.Range(po.Offset(0, 0).Address + ":" + po.Offset(maxX - 2, maxY - 1).Address)

    Dim myChart As Object: Set myChart = dst.Shapes.AddChart.Chart
    With myChart
        .ChartType = xlRadar
        .SetSourceData Source:=range2
        
        With .Legend
            .Position = xlTop
        End With
        
        .Axes(xlValue).TickLabels.NumberFormatLocal = "0_ "
    End With
    With myChart.Parent
         .height = 300  ' resize
         .Width = 300   ' resize
         .Top = 0       ' reposition
         .Left = 0      ' reposition
    End With

    myChart.Parent.Cut
    dst.Select
    po.Offset(3, 0).Select
    dst.Pictures.Paste.Select
    
    pt.ClearTable
    
    s.CurRePo = po.Offset(3 + 23, 0).Address
    
End Function


Private Function mwr(v As Double) As Double
    If v > 360 Then
        mwr = mwrm(v)
    ElseIf v < 0 Then
        mwr = mwra(v)
    Else
        mwr = v
    End If
    
End Function

Private Function mwrm(v As Double) As Double
    v = v - 360
    If v > 360 Then
        v = mwrm(v)
    End If
    
    mwrm = v
End Function

Private Function mwra(v As Double) As Double
    v = v + 360
    If v < 0 Then
        v = mwra(v)
    End If
    
    mwra = v
End Function

Function wr(v As Double) As Integer
    Dim nAngles As Integer: nAngles = 16
    Dim interval As Double: interval = 360 / nAngles
    'Dim arrWr: arrWr = Array("N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE", "S", "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW", "N")

    ', t As String
    'If t = "meteo" Then
    '    v = (90 - v) Mod 360
    'End If
    
    v = mwr(v) + interval / 2
    
    Dim vv As Double: vv = v / interval
    
    wr = Int(vv)

    If wr <> vv Then wr = wr + 1
    If wr = 17 Then wr = 1

End Function
