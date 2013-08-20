Attribute VB_Name = "输出风玫瑰图"

Function 绘制风玫瑰图(s As Object, rst As Object, dst As Object)
    ' 代表年的全年风向、风能频率分布玫瑰图

    s.Pc.Value = "代表年的全年风向、风能频率分布玫瑰图"
    s.Pc = s.Pc.Offset(1, 0)
    
    Dim wds As Object: Set wds = s.Sensors("wd")
    Dim ad: ad = wds.Items
    
    Dim wvs As Object: Set wvs = s.Sensors("wv")
    Dim av: av = wvs.Items
    
    ' 增加数据透视表
    Dim t As Object: Set t = newSheet("tshowwindrose")
    Dim pt As Object: Set pt = newPT(t, s.dataRange, "pt")
    
    For i = 0 To wvs.count - 1
        Dim ssv As Object: Set ssv = av(i)
        Dim ssd As Object: Set ssd = getSSbyH(wds, ssv.height)
        

        showWindrose rst, dst, s, t, pt, ssv, ssd
    Next

    deleteSheet t
End Function


Private Function showWindrose(rst As Object, dst As Object, s As Object, t As Object, pt As Object, ssv As Object, ssd As Object)
    ' 代表年的16个方位扇区不同高度出现风向、风能频率
    Dim arrWr: arrWr = Array("N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE", "S", "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW", "N")

    s.Pc.Value = "CH" + ssv.channel + " " + CStr(ssv.height) + "m 代表年的全年风向、风能频率分布玫瑰图"
    Dim po As Object: Set po = s.Pc.Offset(1, 0)

    pt.ClearTable

    With pt.PivotFields("CH" + ssd.channel + "Wr")
        .Orientation = xlColumnField
        .Position = 1
    End With
    
    Dim ssvn As String: ssvn = "风向频率"
    Dim sspn As String: sspn = "风能频率"
    
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
    
    Dim maxX As Integer: maxX = t.UsedRange.Rows.count
    Dim maxY As Integer: maxY = t.UsedRange.Columns.count
    
    rangeCopy t.Range("A2:" + t.Cells(maxX, maxY).Address), po

    po.Value = ""
    For i = 1 To maxY - 1
        po.Offset(0, i).Value = arrWr(po.Offset(0, i).Value - 1)
    Next
    
    Dim range1 As Object: Set range1 = dst.Range(po.Offset(1, 1), po.Offset(maxX - 2, maxY - 1))
    rangeFV dr:=range1, v:=100, m:=xlMultiply
    range1.NumberFormatLocal = "0.00"


    Dim rangeX As String: rangeX = dst.Name + "!" + po.Offset(0, 1).Address + ":" + po.Offset(0, maxY - 1).Address
    Dim cRangeY As New Collection, cRangeT As New Collection
    For i = 1 To 2
        cRangeY.Add po.Offset(i, 1).Address + ":" + po.Offset(i, maxY - 1).Address
        cRangeT.Add dst.Name + "!" + po.Offset(i, 0).Address
    Next i
    
    drawChart rangeX:=rangeX, cRangeY:=cRangeY, cRangeT:=cRangeT, rst:=dst, dst:=dst, _
            dpo:=po.Offset(3, 0), cType:=xlRadar, height:=300, width:=300, axisFormatY:="0"

    s.Pc = po.Offset(3 + 23, 0)
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
