Attribute VB_Name = "输出"

Sub 生成报表()
    系统初始化
    
    For Each k In Stations
        Dim s As Object: Set s = Stations(k)
        
        计算平均风速风能
        计算风速风能频率
        绘制风玫瑰图
        绘制威布尔曲线
    Next
    
End Sub

Function initCalResult(s As Object)
    calAirDensity s ' 计算空气密度
    
    Dim rst As Object: Set rst = Sheets(s.Sheet1h)
    
    Dim dst As Object: Set dst = Sheets.Add(after:=Sheets(Sheets.Count))
    dst.Name = "result" + s.id
    s.Result = dst.Name

    Dim wvs As Object: Set wvs = s.Sensors("wv")
    Dim a: a = wvs.Items
    Dim ss As Object
    
    ' 增加必要的数据列
    Dim maxX As Integer: maxX = rst.UsedRange.Rows.Count
    Dim maxY As Integer: maxY = rst.UsedRange.Columns.Count
    
        'month
    Dim cellM As Object: Set cellM = rst.Cells(1, maxY + 1)
    Dim cellMF As Object: Set cellMF = cellM.Offset(1, 0)
    Dim cellML As Object: Set cellML = cellM.Offset(maxX - 1, 0)
    cellM.Value = "Month"
    cellMF.Formula = "=month(A2)"
    cellMF.AutoFill Destination:=rst.Range(cellMF.Address + ":" + cellML.Address)
        'hour
    maxY = rst.UsedRange.Columns.Count
    Dim cellH As Object: Set cellH = rst.Cells(1, maxY + 1)
    Dim cellHF As Object: Set cellHF = cellH.Offset(1, 0)
    Dim cellHL As Object: Set cellHL = cellH.Offset(maxX - 1, 0)
    cellH.Value = "Hour"
    cellHF.Formula = "=hour(A2)"
    cellHF.AutoFill Destination:=rst.Range(cellHF.Address + ":" + cellHL.Address)
        'wp
    maxY = rst.UsedRange.Columns.Count
    For j = 0 To wvs.Count - 1
        Set ss = a(j)
        
        Dim cellWp As Object: Set cellWp = rst.Cells(1, maxY + 1 + j)
        Dim cellWpF As Object: Set cellWpF = cellWp.Offset(1, 0)
        Dim cellWpL As Object: Set cellWpL = cellWp.Offset(maxX - 1, 0)
        cellWp.Value = "CH" + ss.channel + "WP"
        cellWpF.Formula = "=" + CStr(s.AirDensity) + "*POWER(" + Replace(rst.Cells(2, ss.avg).Address, "$", "") + ",3)/2" ' =10*POWER(B2,3)/2
        cellWpF.AutoFill Destination:=rst.Range(cellWpF.Address + ":" + cellWpL.Address)
    Next
        'wfre
    maxY = rst.UsedRange.Columns.Count
    For j = 0 To wvs.Count - 1
        Set ss = a(j)
        
        Dim cellWfv As Object: Set cellWfv = rst.Cells(1, maxY + 1 + j)
        Dim cellWfvF As Object: Set cellWfvF = cellWfv.Offset(1, 0)
        Dim cellWfvL As Object: Set cellWfvL = cellWfv.Offset(maxX - 1, 0)
        Dim wfvA As String: wfvA = Replace(rst.Cells(2, ss.avg).Address, "$", "")
        cellWfv.Value = "CH" + ss.channel + "Wfv"
        cellWfvF.Formula = "=IF(" + wfvA + "<=0.5,0.5, ROUND(" + wfvA + ",0))"
        cellWfvF.AutoFill Destination:=rst.Range(cellWfvF.Address + ":" + cellWfvL.Address)
    Next
        ' weibull
    maxY = rst.UsedRange.Columns.Count
    For j = 0 To wvs.Count - 1
        Set ss = a(j)
        
        Dim cellWb As Object: Set cellWb = rst.Cells(1, maxY + 1 + j)
        Dim cellWbF As Object: Set cellWbF = cellWb.Offset(1, 0)
        Dim cellWbL As Object: Set cellWbL = cellWb.Offset(maxX - 1, 0)
        Dim wbA As String: wbA = Replace(rst.Cells(2, ss.avg).Address, "$", "")
        cellWb.Value = "CH" + ss.channel + "Wb"
        cellWbF.Formula = "=IF(" + wbA + "<=0,1, ceiling(" + wbA + ",1))"
        cellWbF.AutoFill Destination:=rst.Range(cellWbF.Address + ":" + cellWbL.Address)
    Next
        'windrose
    maxY = rst.UsedRange.Columns.Count
    Dim wds As Object: Set wds = s.Sensors("wd")
    Dim ad: ad = wds.Items
    Dim ssd As Object
    For j = 0 To wds.Count - 1
        Set ssd = ad(j)
        
        Dim cellWr As Object: Set cellWr = rst.Cells(1, maxY + 1 + j)
        cellWr.Value = "CH" + ssd.channel + "Wr"
        
        For i = 2 To maxX
            rst.Cells(i, maxY + 1 + j).Value = wr(rst.Cells(i, ssd.avg).Value)
        Next
    Next
    
    maxY = rst.UsedRange.Columns.Count
    
    Dim cellL As Object: Set cellL = rst.Cells(maxX, maxY)
    
    s.AddedData = cellM.Address + ":" + cellL.Address
    s.DataRange = rst.Name + "!A1:" + cellL.Address
    
    ' 首行
    Dim pc As Object: Set pc = dst.Range(s.CurRePo)
    pc.Value = "数据日期: " + Format(s.StartTime, "yyyy年mm月dd日") + _
        "～" + Format(s.EndTime, "yyyy年mm月dd日")
    s.CurRePo = pc.Offset(2, 0).Address
    
End Function


' 从传感器集合中选取指定高度的一个传感器
' >>>需要增加含有多个相同高度时的判别方法：例如相关性等
Function getSSbyH(sss As Scripting.Dictionary, h As Double) As Object
    Dim ss As Object
    For Each k In sss
        Dim sst As Object: Set sst = sss(k)
        If sst.height = h Then
            Set getSSbyH = sst
            Exit Function
        End If
        
        If ss Is Nothing Then
            Set ss = sst
        ElseIf Abs(ss.height - height) > Abs(sst.height - height) Then
            Set ss = sst
        End If
    Next

    Set getSSbyH = ss
End Function

' 计算空气密度
Private Function calAirDensity(s As Object)
    s.AirDensity = oConfig.AirDensity
    If oConfig.UseSetAD = "user" Then
        Exit Function
    End If

    Dim ts As Object: Set ts = s.Sensors("t")
    Dim ps As Object: Set ps = s.Sensors("p")
    
    Dim st As Object: Set st = Sheets(s.Sheet1h)
    
    Dim p, t
    If ts.Count > 0 Then
        Dim ta: ta = ts.Items
        Dim sst As Object: Set sst = ta(0)
        Dim ranget As Object: Set ranget = st.Range(arrCol(sst.channel))
        t = Application.WorksheetFunction.Average(ranget)
        
        If ps.Count > 0 Then
            ' 当同时有气温和气压数据时
            Dim pa: pa = ps.Items
            Dim ssp As Object: Set ssp = pa(0)
            Dim rangep As Object: Set rangep = st.Range(arrCol(ssp.channel))
            p = Application.WorksheetFunction.Average(rangep)
            
            s.AirDensity = p * 1000 / (287 * (t + 273))
        Else
            ' 当只有气温数据时
            s.AirDensity = (353.05 / t) * Exp((-0.034) * (height / (t + 273)))
        End If
    End If
End Function
