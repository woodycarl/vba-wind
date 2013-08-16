Attribute VB_Name = "输出湍流强度"

Sub 计算湍流强度()
    系统初始化

    For Each k In Stations
        Dim s As Object: Set s = Stations(k)
        
        If s.CurRePo = "A1" Then
            initCalResult s
        End If
        
        Dim rst As Object: Set rst = Sheets(s.Sheet1h)
        Dim dst As Object: Set dst = Sheets(s.Result)

        Dim wts As New Collection
        
        ' 代表年的不同高度湍流强度
        Dim pc As Object: Set pc = dst.Range(s.CurRePo)
        pc.Value = "代表年的不同高度湍流强度"
        s.CurRePo = pc.Offset(1, 0).Address
        
        Dim wvs As Object: Set wvs = s.Sensors("wv")
        Dim a: a = wvs.Items
        For j = 0 To wvs.count - 1
            Dim ss As Object: Set ss = a(j)
            
            Dim v As WT: Set v = New WT
            With v
                .height = ss.height
                Set .rt60 = rst.Columns(1)
                Set .r60 = rst.Columns(ss.avg)
                Set .rs60 = rst.Columns(ss.Sd)
                .c60 = rst.UsedRange.Rows.count
            End With
            
            If sheetExist(s.Sheet10m) Then
                Dim rst10 As Object: Set rst10 = Sheets(s.Sheet10m)
                Set v.rt10 = rst10.Columns(1)
                Set v.r10 = rst10.Columns(ss.avg)
                Set v.rs10 = rst10.Columns(ss.Sd)
                v.c10 = rst10.UsedRange.Rows.count
            End If

            wts.Add v
            
        Next j
        
        Set pc = dst.Range(s.CurRePo)
        
        calTurbs dst, pc, wts
        
        s.CurRePo = pc.Offset(wvs.count + 17, 0).Address
    Next
    
End Sub

Sub 计算选定多列的湍流强度()
    系统初始化
    Dim wts As New Collection
    
    Dim index As Integer: index = 1
    While True
        
        Dim th: th = InputBox("输入序列" + CStr(index) + "的高度:")
        If th = "" Then
            GoTo ew
        End If

        Dim rt10 As Object: Set rt10 = GetRange("选择序列" + CStr(index) + " 10分钟时间数据:")
        If rt10 Is Nothing Then
            GoTo ew
        End If
        
        Dim c10 As Double: c10 = Sheets(rt10.Parent.Name).UsedRange.Rows.count

        Dim r10 As Object: Set r10 = GetRange("选择序列" + CStr(index) + " 10分钟Avg数据:")
        If r10 Is Nothing Then
            GoTo ew
        End If
        Dim rs10 As Object: Set rs10 = GetRange("选择序列" + CStr(index) + " 10分钟SD数据:")
        If rs10 Is Nothing Then
            GoTo ew
        End If
        
        Dim rt60 As Object: Set rt60 = GetRange("选择序列" + CStr(index) + " 60分钟时间数据:")
        If rt60 Is Nothing Then
            GoTo ew
        End If
        Dim c60 As Double: c60 = Sheets(rt60.Parent.Name).UsedRange.Rows.count
        Dim r60 As Object: Set r60 = GetRange("选择序列" + CStr(index) + " 60分钟Avg数据:")
        If r60 Is Nothing Then
            GoTo ew
        End If

        Dim tn As String: tn = "HH"
        Dim cn As String: cn = "Wt"

        Dim v As WT: Set v = New WT
        With v
            .height = th
            Set .rt10 = rt10
            Set .rt60 = rt60
            Set .r10 = r10
            Set .r60 = r60
            Set .rs10 = rs10
            .c60 = c60
            .c10 = c10
        End With
        wts.Add v

        index = index + 1
    Wend
ew:

    Dim wbn As Object: Set wbn = Workbooks.Add
    Dim dst As Object: Set dst = wbn.Sheets(1)
    Dim dr As Object: Set dr = dst.Range("A1")

    calTurbs dst, dr, wts
End Sub


Function calTurbs(dst As Object, dr As Object, wts As Collection)
    Dim interval As Double: interval = 1
    Dim limit As Double: limit = 0.5

    dr.Value = "测风高度"
    dr.Offset(0, 1).Value = "湍流强度(全部数据)"
    dr.Offset(0, 2).Value = "湍流强度(V=15±0.5m/s)"
    
    dr.Offset(2 + wts.count, 0).Value = "IEC A类"
    dr.Offset(3 + wts.count, 0).Value = "IEC B类"
    dr.Offset(4 + wts.count, 0).Value = "IEC C类"

    For i = 1 To wts.count
        dr.Offset(i, 0).Value = CStr(wts(i).height) + " m"
        Dim r As Scripting.Dictionary
        
        Set r = calTurb(wts(i), interval, limit)

        dr.Offset(i, 1).Value = r("0")
        dr.Offset(i, 2).Value = r("15")
        
        Dim a:  a = r.Items
        Dim k:  k = r.Keys

        dr.Offset(i + 4 + wts.count, 0).Value = wts(i).height
        For j = 1 To r.count - 1
            If a(j) <> 0 Then
                dr.Offset(i + 4 + wts.count, j).Value = CStr(a(j))
                
                If dr.Offset(1 + wts.count, j).Value = "" Then
                    dr.Offset(1 + wts.count, j).Value = k(j)
                    dr.Offset(i + 1 + wts.count, j).Value = 0.75 * 0.16 + 5.6 * 0.16 / dr.Offset(1 + wts.count, j).Value
                    dr.Offset(i + 2 + wts.count, j).Value = 0.75 * 0.14 + 5.6 * 0.14 / dr.Offset(1 + wts.count, j).Value
                    dr.Offset(i + 3 + wts.count, j).Value = 0.75 * 0.12 + 5.6 * 0.12 / dr.Offset(1 + wts.count, j).Value
                End If
            End If
        Next j
        
        dr.Offset(i + 4 + wts.count, 0).Value = CStr(wts(i).height) + " m"
    Next i

    dst.Range(dr.Offset(1, 1).Address + ":" + dr.Offset(wts.count, 2).Address).NumberFormatLocal = "0.00"
    
    Dim maxY As Integer: maxY = 1
    While dr.Offset(1 + wts.count, maxY + 1).Value <> ""
        maxY = maxY + 1
    Wend
    
    Dim range2 As Object: Set range2 = dst.Range(dr.Offset(2 + wts.count, 0).Address + ":" + dr.Offset(4 + 2 * wts.count, maxY).Address)
    Dim range3 As String: range3 = dr.Offset(1 + wts.count, 1).Address + ":" + dr.Offset(1 + wts.count, maxY).Address
    Dim range4 As Object: Set range4 = dst.Range(dr.Offset(1 + wts.count, 0).Address + ":" + dr.Offset(4 + 2 * wts.count, maxY).Address)

    Dim myChart As Object: Set myChart = dst.Shapes.AddChart.Chart
    With myChart
        .ChartType = xlXYScatterSmoothNoMarkers ' xlLine
        .SetSourceData Source:=range4
        '.SeriesCollection(1).XValues = "=" + dst.Name + "!" + range3
        
        With .Legend
            .Position = xlTop
        End With
        
        .SetElement (msoElementPrimaryValueAxisTitleRotated)
        .Axes(xlValue).TickLabels.NumberFormatLocal = "0.0_ "
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "湍流强度"
        .Axes(xlCategory).HasTitle = True
        With .Axes(xlCategory).AxisTitle
            .Format.TextFrame2.TextRange.Characters.Text = "风速 (m/s)"
        End With
        

        With .SeriesCollection(1).Format.Line
            .ForeColor.ObjectThemeColor = msoThemeColorText1
        End With
        With .SeriesCollection(2).Format.Line
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Transparency = 0.3
        End With
        With .SeriesCollection(3).Format.Line
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Transparency = 0.6
        End With
        

    End With
    With myChart.Parent
         .height = 200  ' resize
         .Width = 550   ' resize
         .Top = 0       ' reposition
         .Left = 0      ' reposition
    End With

    myChart.Parent.Cut
    dst.Select
    dr.Offset(1 + wts.count, 0).Select
    dst.Pictures.Paste.Select
    'dst.Paste
    
End Function

Function calTurb(v As WT, interval As Double, limit As Double) As Scripting.Dictionary
    Dim wtn As String: wtn = "Wt"
    Dim tn As String: tn = "Time"
    Dim an As String: an = "Avg"
    oWB.Activate
    
    Set calTurb = CreateObject("Scripting.Dictionary")
    
    Dim t As Object
    Dim tt As Object: Set tt = newSheet("TT")
    
    If v.c10 > 0 Then
        ' 10m数据表
        Dim t10 As Object: Set t10 = newSheet("T10")
    
        Dim rt10a As String: rt10a = v.rt10.Parent.Name + "!" + Replace(v.rt10.Cells(2, 1).Address, "$", "")
        t10.Cells(2, 1).Formula = "=year(" + rt10a + ") & month(" + rt10a + ") & day(" + rt10a + ") & hour(" + rt10a + ")"
        t10.Cells(2, 1).AutoFill Destination:=t10.Range("A2:A" & v.c10)
        
        rangeCopy v.r10, t10.Cells(1, 2)
    
        rangeCopy v.rs10, t10.Cells(1, 3)
        v.r10.Copy
        t10.Cells(1, 3).PasteSpecial Paste:=xlPasteAll, Operation:=xlDivide, SkipBlanks _
            :=False, Transpose:=False
    
        t10.Cells(1, 1).Value = tn
        t10.Cells(1, 2).Value = an
        t10.Cells(1, 3).Value = wtn
        
        Dim tt10 As Object: Set tt10 = newSheet("TT10")
        
        Dim dataRange As String: dataRange = t10.Name + "!" + t10.UsedRange.Address
        Dim pt As Object: Set pt = newPT(tt10, dataRange, "pt")

        With pt.PivotFields(tn)
            .Orientation = xlRowField
            .Position = 1
        End With
    
        pt.AddDataField pt.PivotFields(an), "an", xlAverage
        pt.AddDataField pt.PivotFields(wtn), "wtn", xlMax

        With pt.DataPivotField
            .Orientation = xlColumnField
            .Position = 1
        End With
        
        With pt
            .ColumnGrand = False
            .RowGrand = False
        End With
        
        Set t = newSheet("Twt")
        rangeCopy tt10.UsedRange, t.Cells(1, 1)
        
        Application.DisplayAlerts = False
        tt10.Delete
        t10.Delete
        Application.DisplayAlerts = True

    Else
        ' 60m数据表
        Dim t60 As Object: Set t60 = newSheet("T60")

        Dim rt60a As String: rt60a = v.rt60.Parent.Name + "!" + Replace(v.rt60.Cells(2, 1).Address, "$", "")
        t60.Cells(2, 1).Formula = "=year(" + rt60a + ") & month(" + rt60a + ") & day(" + rt60a + ") & hour(" + rt60a + ")"
        t60.Cells(2, 1).AutoFill Destination:=t60.Range("A2:A" & v.c60)
    
        rangeCopy v.r60, t60.Cells(1, 2)
        
        rangeCopy v.rs60, t60.Cells(1, 3)
        v.r60.Copy
        t60.Cells(1, 3).PasteSpecial Paste:=xlPasteAll, Operation:=xlDivide, SkipBlanks _
            :=False, Transpose:=False

        t60.Cells(1, 1).Value = tn
        t60.Cells(1, 2).Value = an
        t60.Cells(1, 3).Value = wtn
        
        Set t = t60
        t.Name = "Twt"
    End If

    calTurb.Add "0", Application.WorksheetFunction.Average(t.Range("C:C"))

    Dim ra As Object: Set ra = t.Range("B:B")
    Dim maxA As Integer: maxA = Application.WorksheetFunction.Max(ra) 'Application.WorksheetFunction.Ceiling(, 1)
    
    Dim i As Double
    For i = 3 To maxA + interval Step interval 'interval
        calTurb.Add CStr(i), calWt(t, tt, i, limit)
        
        If i < 15 And i + interval > 15 Then
            calTurb.Add CStr(15), calWt(t, tt, 15, limit)
        End If
    Next i

    Application.DisplayAlerts = False
    t.Delete
    tt.Delete
    Application.DisplayAlerts = True

End Function

Function calWt(t As Object, tt As Object, i As Double, limit As Double) As Double
    tt.UsedRange.Clear
    t.UsedRange.AutoFilter Field:=2, Criteria1:=">" & (i - limit), _
        Operator:=xlAnd, Criteria2:="<" & (i + limit)
    rangeCopy t.UsedRange, tt.Range("A1")
    
    If tt.UsedRange.Rows.count < 2 Then
        calWt = 0
        Exit Function
    End If
    
    calWt = Application.WorksheetFunction.Average(tt.Columns(3))
End Function

Function GetRange(str As String) As Object
    On Error Resume Next
    Set GetRange = Application.InputBox(prompt:=str, Type:=8)
End Function
