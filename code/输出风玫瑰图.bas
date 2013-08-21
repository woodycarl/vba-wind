Attribute VB_Name = "输出风玫瑰图"
Private arrWr

Function 绘制风玫瑰图(s As Object, rst As Object, dst As Object)
    ' 代表年的全年风向、风能频率分布玫瑰图

    arrWr = Array("N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE", "S", "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW", "N")

    s.Pc.Value = "代表年的全年风向、风能频率分布玫瑰图"
    s.Pc = s.Pc.Offset(1, 0)
    
    Dim wds As Object: Set wds = s.sensors("风向")
    Dim ad: ad = wds.Items
    
    Dim wvs As Object: Set wvs = s.sensors("风速")
    Dim av: av = wvs.Items
    
    ' 增加数据透视表
    Dim t As Object: Set t = newSheet("tshowwindrose")
    Dim pt As Object: Set pt = newPT(t, s.dataRange, "pt")
    
    For i = 0 To wvs.count - 1
        Dim ssv As Object: Set ssv = av(i)
        Dim ssd As Object: Set ssd = getSSbyH(wds, ssv.height)
        
        showWindrose rst:=rst, dst:=dst, s:=s, t:=t, pt:=pt, ssv:=ssv, ssd:=ssd, _
            title:="CH" + ssv.channel + " " + CStr(ssv.height) + "m 代表年的全年风向、风能频率分布玫瑰图"
        
        showWindroseMs rst:=rst, dst:=dst, s:=s, ssv:=ssv, ssd:=ssd
    Next

    deleteSheet t
End Function

Private Function showWindrose(rst As Object, dst As Object, s As Object, t As Object, pt As Object, _
        ssv As Object, ssd As Object, title As String, Optional cTitle As String = "")
    ' 代表年的16个方位扇区不同高度出现风向、风能频率
    s.Pc.Value = title

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
    
    Dim cLegend As Variant
    If cTitle = "" Then
        cLegend = xlTop
    Else
        cLegend = -2
    End If
    drawChart rangeX:=rangeX, cRangeY:=cRangeY, cRangeT:=cRangeT, rst:=dst, dst:=dst, _
            dpo:=po.Offset(3, 0), cType:=xlRadar, height:=300, width:=300, axisFormatY:="0", _
            cTitle:=cTitle, cLegend:=cLegend

    s.Pc = s.Pc.Offset(0, 19)
End Function

Private Function showWindroseMs(rst As Object, dst As Object, s As Object, _
        ssv As Object, ssd As Object)
    Dim tt As Object: Set tt = newSheet("tshowWindroseMs")
    Dim ptt As Object: Set ptt = newPT(tt, rst.UsedRange.Address, "pttms")

    With ptt.PivotFields("Month")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    Dim colNum As Integer: colNum = getColNum(rst, "Month")


    With ptt.PivotFields("Month")
        Dim i As Integer
        For i = 1 To .PivotItems.count
            If .PivotItems(i).Name <> "" Then
                Dim trst As Object: Set trst = newSheet("trst")
                
                rst.UsedRange.AutoFilter
                rst.UsedRange.AutoFilter Field:=colNum, Criteria1:="=" & .PivotItems(i).Name, Operator:=xlAnd
                
                rangeCopy rst.UsedRange, trst.Cells(1, 1)
                trst.Columns(1).NumberFormatLocal = "yyyy/m/d hh:mm"
                
                
                ' 增加数据透视表
                Dim t As Object: Set t = newSheet("tshowWindroseMst")
                Dim pt As Object: Set pt = newPT(t, trst.Name + "!" + trst.UsedRange.Address, "pt")
                
                showWindrose rst:=trst, dst:=dst, s:=s, t:=t, pt:=pt, ssv:=ssv, ssd:=ssd, title:="CH" + ssv.channel, _
                    cTitle:=.PivotItems(i).Name + "月"
                
                
                rst.UsedRange.AutoFilter
                deleteSheet trst
                deleteSheet t
            End If
        Next i
    End With

    s.Pc = s.Pc.Offset(27, 1 - s.Pc.Column)

    deleteSheet tt
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
