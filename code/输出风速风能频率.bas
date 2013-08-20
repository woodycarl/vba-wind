Attribute VB_Name = "输出风速风能频率"

Sub 计算风速风能频率(s As Object, rst As Object, dst As Object)



    ' 不同高度风速和风能频率分布
    s.Pc.Value = "不同高度风速和风能频率分布"
    s.Pc = s.Pc.Offset(1, 0)
    
    Dim wvs As Object: Set wvs = s.Sensors("wv")
    Dim a: a = wvs.Items
    For j = 0 To wvs.count - 1
        Dim ss As Object: Set ss = a(j)
        
        showWfv rst:=rst, dst:=dst, s:=s, ss:=ss
    Next


End Sub

Private Function showWfv(rst As Object, dst As Object, s As Object, ss As Object)
    ' 代表年不同高度风速和风能频率分布
    s.Pc.Value = "CH" + ss.channel + " " + CStr(ss.height) + "m 高度代表年风速和风能频率分布直方图"
    Dim po As Object: Set po = s.Pc.Offset(1, 0)

    ' 增加数据透视表
    Dim t As Object: Set t = newSheet("tcalwvpfr")
    Dim pt As Object: Set pt = newPT(t, s.dataRange, "pt")

    Dim wfvn As String: wfvn = "CH" + ss.channel + "Wfv"
    With pt.PivotFields(wfvn)
        .Orientation = xlRowField
        .Position = 1
    End With

    Dim wvn As String: wvn = "风速频率"
    Dim wpn As String: wpn = "风功率频率"
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
    
    
    Dim maxX As Integer: maxX = t.UsedRange.Rows.count
    Dim maxY As Integer: maxY = t.UsedRange.Columns.count
    rangeCopy t.UsedRange, po
    
    If po.Offset(1, 0).Value = 0.5 Then
        po.Offset(1, 0).Value = "≤0.5"
    End If

    Dim rangeData As Object: Set rangeData = dst.Range(po.Offset(1, 1), po.Offset(maxX - 1, maxY - 1))
    rangeFV rangeData, 100, xlMultiply
    rangeData.NumberFormatLocal = "0.00"

    po.Value = "风速区间(m/s)"
    Dim range1 As Object: Set range1 = dst.Range(po.Offset(1, 0), po.Offset(maxX - 1, 0))
    range1.NumberFormatLocal = "0" '.0

    Dim rangeX As String: rangeX = dst.Name + "!" + po.Offset(1, 0).Address + ":" + po.Offset(maxX - 1, 0).Address
    Dim cRangeY As New Collection, cRangeT As New Collection
    For i = 1 To 2
        cRangeY.Add po.Offset(1, i).Address + ":" + po.Offset(maxX - 1, i).Address
        cRangeT.Add dst.Name + "!" + po.Offset(0, i).Address
    Next i
    
    drawChart rangeX:=rangeX, cRangeY:=cRangeY, cRangeT:=cRangeT, rst:=dst, dst:=dst, _
            dpo:=po.Offset(maxX, 0), cType:=xlColumnClustered, axisTitleX:="风速 (m/s)", _
            axisTitleY:="频率 (%)"

    s.Pc = po.Offset(maxX + 16, 0)

    ' 删除增加的表
    deleteSheet t
End Function

