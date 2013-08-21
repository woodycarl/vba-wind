Attribute VB_Name = "输出平均风速风能"
Private unitDic As Scripting.Dictionary

Function 计算平均风速风能(s As Object, rst As Object, dst As Object)
    init

    ' 增加数据透视表
    Dim t As Object: Set t = newSheet("tcalavg")
    Dim pt As Object: Set pt = newPT(t, s.dataRange, "pt")

    ' 1、代表年不同高度月平均风速
    s.Pc.Value = "代表年不同高度月平均风速"
    showAvg rst:=rst, dst:=dst, t:=t, s:=s, pt:=pt, cat:="Avg"

    ' 2、代表年不同高度月平均风功率密度
    s.Pc.Value = "代表年不同高度月平均风功率密度"
    showAvg rst:=rst, dst:=dst, t:=t, s:=s, pt:=pt, cat:="WP"

    ' 3、代表年不同高度小时平均风速
    s.Pc.Value = "代表年不同高度小时平均风速"
    showAvgH rst:=rst, dst:=dst, t:=t, s:=s, pt:=pt, cat:="Avg"

    ' 4、代表年不同高度小时平均风功率密度
    s.Pc.Value = "代表年不同高度小时平均风功率密度"
    showAvgH rst:=rst, dst:=dst, t:=t, s:=s, pt:=pt, cat:="WP"
    
    ' 5、代表年的各月风速风功率日变化曲线图
    s.Pc.Value = "代表年的各月风速风功率日变化曲线图"
    s.Pc = s.Pc.Offset(1, 0)
    showAvgMHs rst:=rst, dst:=dst, s:=s
    
    deleteSheet t
End Function

Private Function init()
    Set unitDic = CreateObject("Scripting.Dictionary")
    unitDic.Add "Avg", "风速 (m/s)"
    unitDic.Add "WP", "风功率密度 (W/m2)"
End Function

Private Function showAvg(rst As Object, dst As Object, t As Object, s As Object, pt As Object, cat As String)
    ' 代表年不同高度月平均
    Dim unit As String: unit = unitDic(cat)
    Dim po As Object: Set po = s.Pc

    pt.ClearTable
    
    With pt.PivotFields("Month")
        .Orientation = xlColumnField
        .Position = 1
    End With

    Dim wvs As Object: Set wvs = s.sensors("风速")
    Dim a: a = wvs.Items
    For j = 0 To wvs.count - 1
        Dim ss As Object: Set ss = a(j)
        Dim ssn As String: ssn = "CH" + ss.channel + cat

        pt.AddDataField pt.PivotFields(ssn), ss.channel + " " + CStr(ss.height) + "m", xlAverage
    Next
    
    If wvs.count > 1 Then
        With pt.DataPivotField
            .Orientation = xlRowField
            .Position = 1
        End With
    End If
    
    Dim maxX As Integer: maxX = t.UsedRange.Rows.count
    Dim maxY As Integer: maxY = t.UsedRange.Columns.count
    rangeCopy t.Range(t.Cells(2, 1), t.Cells(maxX, maxY)), po.Offset(1, 1)
    
    With dst.Range(po.Offset(2, 2), po.Offset(maxX, maxY + 1))
        .NumberFormatLocal = "0.00"
    End With

    Dim rangeX As String: rangeX = dst.Name + "!" + po.Offset(1, 2).Address + ":" + po.Offset(1, maxY - 1).Address
    Dim cRangeY As New Collection, cRangeT As New Collection
    For i = 1 To wvs.count
        cRangeY.Add po.Offset(i + 1, 2).Address + ":" + po.Offset(i + 1, maxY - 1).Address
        cRangeT.Add dst.Name + "!" + po.Offset(i + 1, 1).Address
    Next i
    
    Dim cType As Variant
    If maxY = 3 Then
        cType = xlColumnClustered
    Else
        cType = xlLine
    End If
    
    drawChart rangeX:=rangeX, cRangeY:=cRangeY, cRangeT:=cRangeT, rst:=dst, dst:=dst, _
            dpo:=po.Offset(2 + wvs.count, 0), cType:=cType, axisTitleX:="月份", axisTitleY:=unit, _
            axisFormatY:="0.0"
    
    rangeMerge dst.Range(po.Offset(1, 0), po.Offset(1, 1)), "时间 (月)"
    rangeMerge dst.Range(po.Offset(2, 0), po.Offset(wvs.count + 1, 0)), unit
    po.Offset(1, maxY).Value = "平均"
    
    pt.ClearTable
    s.Pc = po.Offset(2 + wvs.count + 16, 0)
End Function



Private Function showAvgH(rst As Object, dst As Object, t As Object, s As Object, pt As Object, cat As String)
    ' 代表年不同高度小时平均
    Dim unit As String: unit = unitDic(cat)
    Dim po As Object: Set po = s.Pc
    
    pt.ClearTable

    With pt.PivotFields("Hour")
        .Orientation = xlRowField
        .Position = 1
    End With

    Dim wvs As Object: Set wvs = s.sensors("风速")
    Dim a: a = wvs.Items
    For j = 0 To wvs.count - 1
        Dim ss As Object: Set ss = a(j)
        Dim ssn As String: ssn = "CH" + ss.channel + cat

        pt.AddDataField pt.PivotFields(ssn), ss.channel + " " + CStr(ss.height) + "m", xlAverage
    Next
    
    If wvs.count > 1 Then
        With pt.DataPivotField
            .Orientation = xlColumnField
            .Position = 1
        End With
    End If

    Dim maxX As Integer: maxX = t.UsedRange.Rows.count
    Dim maxY As Integer: maxY = t.UsedRange.Columns.count
    rangeCopy t.UsedRange, po.Offset(2, 0)
    
    With dst.Range(po.Offset(3, 1), po.Offset(maxX + 1, maxY - 1))
        .NumberFormatLocal = "0.00"
    End With


    Dim rangeX As String: rangeX = dst.Name + "!" + po.Offset(3, 0).Address + ":" + po.Offset(maxX, 0).Address
    Dim cRangeY As New Collection, cRangeT As New Collection
    For i = 1 To wvs.count
        cRangeY.Add po.Offset(3, i).Address + ":" + po.Offset(maxX, i).Address
        cRangeT.Add dst.Name + "!" + po.Offset(2, i).Address
    Next i
    
    drawChart rangeX:=rangeX, cRangeY:=cRangeY, cRangeT:=cRangeT, rst:=dst, dst:=dst, _
            dpo:=po.Offset(2 + maxX, 0), axisTitleX:="小时", axisTitleY:=unit, _
            axisFormatY:="0.0"


    rangeMerge dst.Range(po.Offset(1, 0), po.Offset(2, 0)), "时间 (小时)"
    rangeMerge dst.Range(po.Offset(1, 1), po.Offset(1, wvs.count)), unit
    po.Offset(maxX + 1, 0).Value = "平均"

    pt.ClearTable
    s.Pc = po.Offset(28 + 16, 0)
End Function

' 当数据不足时，需要增加判断,错误处理
Private Function showAvgMHs(rst As Object, dst As Object, s As Object)

    Dim t As Object: Set t = newSheet("tshowavgmhs")
    Dim pt As Object: Set pt = newPT(t, rst.UsedRange.Address, "ptmhs")

    With pt.PivotFields("Month")
        .Orientation = xlRowField
        .Position = 1
    End With

    Dim wvs As Object: Set wvs = s.sensors("风速")
    Dim a: a = wvs.Items
    For j = 0 To wvs.count - 1
        Dim ss As Object: Set ss = a(j)
        s.Pc.Value = "CH" + ss.channel
        s.Pc = s.Pc.Offset(1, 0)
        
        With pt.PivotFields("Month")
            Dim i As Integer
            For i = 1 To .PivotItems.count
                If .PivotItems(i).Name <> "" Then
                    showAvgMH rst, dst, s, ss, .PivotItems(i).Name
                End If
            Next i
        End With

        s.Pc = s.Pc.Offset(28, 1 - s.Pc.Column)
    Next
    
    deleteSheet t
End Function


Private Function showAvgMH(rst As Object, dst As Object, s As Object, ss As Object, _
        month As String)
    ' 代表年不同高度月份小时平均
    Dim po As Object: Set po = s.Pc

    Dim t As Object: Set t = newSheet("tshowavgmh")
    Dim pt As Object: Set pt = newPT(t, rst.UsedRange.Address, "ptmh")

    With pt.PivotFields("Month")
        .Orientation = xlRowField
        .Position = 1
        Dim x As Integer
        For x = 1 To .PivotItems.count
            If .PivotItems(x).Name <> month Then
                .PivotItems(x).Visible = False
            End If
        Next x
    End With
    
    With pt.PivotFields("Hour")
        .Orientation = xlRowField
        .Position = 2
    End With
    
    hidePTsum pt
    
    pt.AddDataField pt.PivotFields("CH" + ss.channel + "Avg"), "风速", xlAverage
    pt.AddDataField pt.PivotFields("CH" + ss.channel + "WP"), "风功率", xlAverage

    With pt.DataPivotField
        .Orientation = xlColumnField
        .Position = 1
    End With

    Dim maxX As Integer: maxX = t.UsedRange.Rows.count
    Dim maxY As Integer: maxY = t.UsedRange.Columns.count
    rangeCopy t.UsedRange, po
    
    With dst.Range(po.Offset(1, 1), po.Offset(maxX - 1, maxY - 1))
        .NumberFormatLocal = "0.00"
    End With

    Dim rangeX As String: rangeX = dst.Name + "!" + po.Offset(2, 0).Address + ":" + po.Offset(maxX - 1, 0).Address
    Dim cRangeY As New Collection, cRangeT As New Collection
    For i = 1 To 2
        cRangeY.Add po.Offset(2, i).Address + ":" + po.Offset(maxX - 1, i).Address
        cRangeT.Add dst.Name + "!" + po.Offset(0, i).Address
    Next i
    Dim secondarySeries As New Collection
    secondarySeries.Add 2
    Dim c As Object
    Set c = drawChart(rangeX:=rangeX, cRangeY:=cRangeY, cRangeT:=cRangeT, rst:=dst, dst:=dst, _
            dpo:=po.Offset(3, 0), axisTitleY:=unitDic("Avg"), height:=280, width:=500, _
            secondaryAxisTitleY:=unitDic("WP"), secondarySeries:=secondarySeries, _
            cTitle:=month + "月", cLegend:=-2, cLTop:=0, cLLeft:=0, cLHeight:=25, cLWidth:=500, _
            axisFormatY:="0.0")
    
    chart2pic myChart:=c, dst:=dst, dpo:=po.Offset(18, 6), resize:=0.4
    
    po.Value = ""
    po.Offset(1, 0).Value = "平均"

    deleteSheet t
    s.Pc = po.Offset(0, 12)
End Function














