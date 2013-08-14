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
        For j = 0 To wvs.Count - 1
            Dim ss As Object: Set ss = a(j)
            
            Dim v As WT: Set v = New WT
            With v
                .height = ss.height
                Set .rt60 = rst.Columns(1)
                Set .r60 = rst.Columns(ss.avg)
                Set .rs60 = rst.Columns(ss.Sd)
                .c60 = rst.UsedRange.Rows.Count
                .cn = "Wt"
                .tn = "HH"
            End With
            
            If sheetExist(s.Sheet10m) Then
                Dim rst10 As Object: Set rst10 = Sheets(s.Sheet10m)
                Set v.rt10 = rst10.Columns(1)
                Set v.r10 = rst10.Columns(ss.avg)
                Set v.rs10 = rst10.Columns(ss.Sd)
                v.c10 = rst10.UsedRange.Rows.Count
            End If

            wts.Add v
            
        Next j
        
        Set pc = dst.Range(s.CurRePo)
        
        calTurbs dst, pc, wts
        
        s.CurRePo = pc.Offset(wvs.Count + 3, 0).Address
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
        
        Dim c10 As Double: c10 = Sheets(rt10.Parent.Name).UsedRange.Rows.Count

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
        Dim c60 As Double: c60 = Sheets(rt60.Parent.Name).UsedRange.Rows.Count
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
            .cn = cn
            .tn = tn
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

    dr.Value = "测风高度"
    dr.Offset(0, 1).Value = "湍流强度(全部数据)"
    dr.Offset(0, 2).Value = "湍流强度(V=15±0.5m/s)"

    For i = 1 To wts.Count
        dr.Offset(i, 0).Value = CStr(wts(i).height) + " m"
        dr.Offset(i, 1).Value = calTurb(wts(i), True)
        dr.Offset(i, 2).Value = calTurb(wts(i), False)
    Next

    dst.Range(dr.Offset(1, 1).Address + ":" + dr.Offset(wts.Count, 2).Address).NumberFormatLocal = "0.00"

End Function

Function calTurb(v As WT, ba As Boolean) As Double
    oWB.Activate
    Dim tn: tn = "Twt"
    Sheets.Add After:=Sheets(Sheets.Count)
    Dim twt As Object: Set twt = ActiveSheet
    twt.Name = tn
 
    twt.Cells(1, 1).Value = v.tn
    twt.Cells(1, 2).Value = v.cn

    For i = 2 To v.c10
        twt.Cells(i, 1).Value = Format(v.rt10.Cells(i, 1).Value, "yyyymdhh")
        twt.Cells(i, 2).Value = v.rs10.Cells(i, 1).Value / v.r10.Cells(i, 1).Value
    Next i

    If ba Then
        calTurb = Application.WorksheetFunction.Average(twt.Range("B:B"))
    Else
        Dim sum As Double
        Dim num As Double
    
        ' 透视表数据
        Dim DataRange As String: DataRange = twt.Name + "!" + twt.UsedRange.Address
        ' 增加数据透视表
        oWB.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
            DataRange, Version:=xlPivotTableVersion14). _
            CreatePivotTable TableDestination:=oTemp.Name + "!R1C1", TableName:="pt", _
            DefaultVersion:=xlPivotTableVersion14
        Dim pt As Object: Set pt = oTemp.PivotTables("pt")
        'oTemp.Select
        pt.ClearTable
    
        With pt.PivotFields(v.tn)
            .Orientation = xlRowField
            .Position = 1
        End With
    
        pt.AddDataField pt.PivotFields(v.cn), v.cn + " " + CStr(v.height) + "m", xlMax

        For i = 2 To v.c60
            If v.r60.Cells(i, 1).Value > 14.5 And v.r60.Cells(i, 1).Value < 15.5 Then
            
                Dim sx As String: sx = Format(v.rt60.Cells(i, 1).Value, "yyyymdh")
                
                pt.PivotFields(v.tn).PivotFilters.Add Type:= _
                    xlCaptionEquals, Value1:=sx
                
                sum = sum + oTemp.Cells(2, 2).Value
                num = num + 1

                pt.PivotFields(v.tn).ClearAllFilters
                
            End If
            
        Next i
        
        ' 清除数据透视表、删除增加的sheet
        oTemp.Range(pt.TableRange2.Address).Delete Shift:=xlUp
    
        calTurb = sum / num
    End If
    
    Application.DisplayAlerts = False
    twt.Delete
    Application.DisplayAlerts = True

End Function

Function GetRange(str As String) As Object
    On Error Resume Next
    Set GetRange = Application.InputBox(prompt:=str, Type:=8)
End Function
