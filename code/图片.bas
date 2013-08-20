Attribute VB_Name = "图片"
' 以表格内容为数据生成图表
'
' @ rangeX: 横坐标range "st!A1:A3"
' @ cRangeY: 纵坐标range集合 "B1:B3"
' @ cRangeT: 标题range集合 "st!D1:D3"
' @ rst: 源数据sheet
' @ dst: 需要生成图片的sheet
' @ dpo: 生成图片所在的位置
'
' @ axisTitleX: Optional 横坐标标题,默认无
' @ axisTitleY: Optional 纵坐标标题,默认无
' @ cType: Optional 图表类型,默认为xlLine. xlColumnClustered,xlLine
' @ cToPic: Optional 是否生成为图片,默认为false
' @ axisFormatY: Optional 纵坐标数字格式
' @ width: Optional 宽,默认550
' @ height: Optional 高,默认200
' @ cLegend: Optional 默认靠上
' @ cLTop: Optional 默认无
' @ cLLeft: Optional 默认无
' @ tickSpacingX: Optional 横坐标标签间隔 默认1
' @ secondaryAxisTitleY: Optional 次要纵坐标
Function drawChart(rangeX As String, _
        cRangeY As Collection, _
        cRangeT As Collection, _
        rst As Object, _
        dst As Object, _
        dpo As Object, _
        Optional axisTitleX As String = "", _
        Optional axisTitleY As String = "", _
        Optional cType As Variant = xlLine, _
        Optional cToPic As Boolean = False, _
        Optional axisFormatX As String = "", _
        Optional axisFormatY As String = "", _
        Optional secondaryAxisFormatY As String = "0.0", _
        Optional width As Double = 550, _
        Optional height As Double = 200, _
        Optional cTitle As String = "", _
        Optional cLegend As Variant = xlTop, _
        Optional cLTop As Double = 0, _
        Optional cLLeft As Double = 0, _
        Optional cLWidth As Double = 100, _
        Optional cLHeight As Double = 20, _
        Optional tickSpacingX As Integer = 1, _
        Optional secondarySeries As Collection, _
        Optional secondaryAxisTitleY As String = "") As Object
    
    Dim myChart As Object: Set myChart = dst.Shapes.AddChart.Chart
    With myChart
        .ChartType = cType

        .SetSourceData Source:=rst.Range(cRangeY(1))
        .SeriesCollection(1).XValues = "=" + rangeX
        .SeriesCollection(1).Name = "=" + cRangeT(1)
        
        If cRangeY.count > 1 Then
            For i = 2 To cRangeY.count
                .SeriesCollection.NewSeries
                .SeriesCollection(i).XValues = "=" + rangeX
                .SeriesCollection(i).Values = "=" + rst.Name + "!" + cRangeY(i)
                .SeriesCollection(i).Name = "=" + cRangeT(i)
            Next i
        End If
        
        With .Legend
            If cLegend = -1 Then
                .Delete
            ElseIf cLegend = -2 Then
                .IncludeInLayout = False
                
                .left = cLLeft
                .top = cLTop
                .width = cLWidth
                .height = cLHeight
            Else
                .Position = cLegend
            End If
        End With
        
        If axisTitleX <> "" Then
            .Axes(xlCategory).HasTitle = True
            With .Axes(xlCategory).AxisTitle
                .Format.TextFrame2.TextRange.Characters.Text = axisTitleX
            End With

        End If
        If axisFormatX <> "" Then
            .Axes(xlCategory).TickLabels.NumberFormatLocal = axisFormatX
        End If
        
        If axisTitleY <> "" Then
            .SetElement (msoElementPrimaryValueAxisTitleRotated)
            .Axes(xlValue, xlPrimary).AxisTitle.Text = axisTitleY

            If secondaryAxisTitleY <> "" Then
                For i = 1 To secondarySeries.count
                    .SeriesCollection(secondarySeries(i)).AxisGroup = 2
                Next i
            
                .SetElement (msoElementSecondaryValueAxisTitleRotated)
                .Axes(xlValue, xlSecondary).AxisTitle.Text = secondaryAxisTitleY
                
                .Axes(xlValue, xlSecondary).TickLabels.NumberFormatLocal = secondaryAxisFormatY
            End If
        End If
        If axisFormatY <> "" Then
            .Axes(xlValue).TickLabels.NumberFormatLocal = axisFormatY
        End If

        If cTitle <> "" Then
            .SetElement (msoElementChartTitleAboveChart)
            With .ChartTitle
                .Text = cTitle
    
                With .Characters.Font
                    .Size = 14
                End With
            End With
        End If

        If tickSpacingX <> 1 Then
            .Axes(xlCategory).TickLabelSpacing = tickSpacingX
            .Axes(xlCategory).TickMarkSpacing = tickSpacingX
        End If
        
    End With
    
    With myChart.Parent
         .height = height
         .width = width
         .top = dpo.top
         .left = dpo.left
    End With

    If cToPic Then
        myChart.Parent.Cut
        dst.Select
        dpo.Select
        dst.Pictures.Paste.Select
        Set drawChart = Nothing
    Else
        Set drawChart = myChart
        'dst.Paste
    End If
End Function

Function addTBtoChart(myChart As Object, content As String, Optional width As Double = 100, Optional height As Double = 100, _
        Optional top As Double = 50, Optional left As Double = 400)
    With myChart.Shapes.AddTextbox(msoTextOrientationHorizontal, left, top, width, height)
        With .TextFrame2.TextRange
            .Characters.Text = content
            With .Characters.Font
                .Size = 11
            End With
        End With
    End With
End Function

Function chart2pic(myChart As Object, dst As Object, dpo As Object, Optional resize As Double = -1)
    myChart.Parent.Copy
    dst.Select
    dpo.Select
    dst.Pictures.Paste.Select
    
    If resize <> -1 Then
        Selection.ShapeRange.ScaleHeight resize, msoFalse, msoScaleFromTopLeft
    End If
End Function

