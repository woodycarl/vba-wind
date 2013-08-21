Attribute VB_Name = "输出"

Sub 生成报表()
    calResult "所有"
End Sub

Sub 显示表1()
    calResult "表1"
End Sub
Sub 显示平均风速风能()
    calResult "平均风速风能"
End Sub
Sub 显示风速风能频率()
    calResult "风速风能频率"
End Sub
Sub 显示风玫瑰图()
    calResult "风玫瑰图"
End Sub
Sub 显示威布尔曲线()
    calResult "威布尔曲线"
End Sub
Sub 显示风切变指数()
    calResult "风切变指数"
End Sub
Sub 显示湍流强度()
    calResult "湍流强度"
End Sub
Function calResult(str As String)
    系统初始化
    
    For Each k In Stations
        Dim s As Object: Set s = Stations(k)
        
        If s.CurRePo = "A1" Then
            initCalResult s
        End If
        
        Dim rst60 As Object: Set rst60 = Sheets(s.Sheet1h)
        Dim dst As Object: Set dst = Sheets(s.Result)
        Dim rst10 As Object
        If sheetExist(s.Sheet10m) Then
            Set rst10 = Sheets(s.Sheet10m)
        Else
            Set rst10 = Nothing
        End If
        
        Select Case str
            Case "表1"
                showInfo s:=s, dst:=dst
            Case "平均风速风能"
                计算平均风速风能 s:=s, rst:=rst60, dst:=dst
            Case "风速风能频率"
                计算风速风能频率 s:=s, rst:=rst60, dst:=dst
            Case "风玫瑰图"
                绘制风玫瑰图 s:=s, rst:=rst60, dst:=dst
            Case "威布尔曲线"
                绘制威布尔曲线 s:=s, rst:=rst60, dst:=dst
            Case "风切变指数"
                计算风切变指数 s:=s, rst:=rst60, dst:=dst
            Case "湍流强度"
                计算湍流强度 s:=s, rst10:=rst10, rst60:=rst60, dst:=dst
            Case "所有"
                showInfo s:=s, dst:=dst
                计算平均风速风能 s:=s, rst:=rst60, dst:=dst
                计算风速风能频率 s:=s, rst:=rst60, dst:=dst
                绘制风玫瑰图 s:=s, rst:=rst60, dst:=dst
                绘制威布尔曲线 s:=s, rst:=rst60, dst:=dst
                计算风切变指数 s:=s, rst:=rst60, dst:=dst
                计算湍流强度 s:=s, rst10:=rst10, rst60:=rst60, dst:=dst
        End Select
        
    Next
End Function

Function initCalResult(s As Object)
    calAirDensity s ' 计算空气密度
    
    Dim rst As Object: Set rst = Sheets(s.Sheet1h)
    
    Dim dst As Object: Set dst = Sheets.Add(After:=Sheets(Sheets.count))
    dst.Name = "result" + s.id
    s.Result = dst.Name

    Dim wvs As Object: Set wvs = s.sensors("风速")
    Dim a: a = wvs.Items
    Dim ss As Object
    
    ' 增加必要的数据列
    Dim maxX As Integer: maxX = rst.UsedRange.Rows.count
    Dim maxY As Integer: maxY = rst.UsedRange.Columns.count
    
        'month
    Dim cellM As Object: Set cellM = rst.Cells(1, maxY + 1)
    Dim cellMF As Object: Set cellMF = cellM.Offset(1, 0)
    Dim cellML As Object: Set cellML = cellM.Offset(maxX - 1, 0)
    cellM.Value = "Month"
    cellMF.Formula = "=month(A2)"
    cellMF.AutoFill Destination:=rst.Range(cellMF.Address + ":" + cellML.Address)
        'hour
    maxY = rst.UsedRange.Columns.count
    Dim cellH As Object: Set cellH = rst.Cells(1, maxY + 1)
    Dim cellHF As Object: Set cellHF = cellH.Offset(1, 0)
    Dim cellHL As Object: Set cellHL = cellH.Offset(maxX - 1, 0)
    cellH.Value = "Hour"
    cellHF.Formula = "=hour(A2)"
    cellHF.AutoFill Destination:=rst.Range(cellHF.Address + ":" + cellHL.Address)
        'wp
    maxY = rst.UsedRange.Columns.count
    For j = 0 To wvs.count - 1
        Set ss = a(j)
        
        Dim cellWp As Object: Set cellWp = rst.Cells(1, maxY + 1 + j)
        Dim cellWpF As Object: Set cellWpF = cellWp.Offset(1, 0)
        Dim cellWpL As Object: Set cellWpL = cellWp.Offset(maxX - 1, 0)
        cellWp.Value = "CH" + ss.channel + "WP"
        cellWpF.Formula = "=" + CStr(s.AirDensity) + "*POWER(" + Replace(rst.Cells(2, ss.avg).Address, "$", "") + ",3)/2" ' =10*POWER(B2,3)/2
        cellWpF.AutoFill Destination:=rst.Range(cellWpF.Address + ":" + cellWpL.Address)
    Next
        'wfre
    maxY = rst.UsedRange.Columns.count
    For j = 0 To wvs.count - 1
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
    maxY = rst.UsedRange.Columns.count
    For j = 0 To wvs.count - 1
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
    maxY = rst.UsedRange.Columns.count
    Dim wds As Object: Set wds = s.sensors("风向")
    Dim ad: ad = wds.Items
    Dim ssd As Object
    For j = 0 To wds.count - 1
        Set ssd = ad(j)
        
        Dim cellWr As Object: Set cellWr = rst.Cells(1, maxY + 1 + j)
        cellWr.Value = "CH" + ssd.channel + "Wr"
        
        For i = 2 To maxX
            rst.Cells(i, maxY + 1 + j).Value = wr(rst.Cells(i, ssd.avg).Value)
        Next
    Next
    
    maxY = rst.UsedRange.Columns.count
    
    Dim cellL As Object: Set cellL = rst.Cells(maxX, maxY)
    
    s.AddedData = cellM.Address + ":" + cellL.Address
    s.dataRange = rst.Name + "!A1:" + cellL.Address
    
    ' 首行
    Dim Pc As Object: Set Pc = dst.Range(s.CurRePo)
    Pc.Value = "数据日期: " + Format(s.StartTime, "yyyy年mm月dd日") + _
        "～" + Format(s.EndTime, "yyyy年mm月dd日")
    s.CurRePo = Pc.Offset(2, 0).Address
    
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
        ElseIf Abs(ss.height - height) < Abs(sst.height - height) Then
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

    Dim ts As Object: Set ts = s.sensors("气温")
    Dim ps As Object: Set ps = s.sensors("气压")
    
    Dim st As Object: Set st = Sheets(s.Sheet1h)
    
    Dim p, t
    If ts.count > 0 Then
        Dim ta: ta = ts.Items
        t = Application.WorksheetFunction.Average(st.Columns(ta(0).avg))
        
        If ps.count > 0 Then
            ' 当同时有气温和气压数据时
            Dim pa: pa = ps.Items
            p = Application.WorksheetFunction.Average(st.Columns(pa(0).avg))
            
            s.AirDensity = p * 1000 / (287 * (t + 273))
        Else
            ' 当只有气温数据时
            s.AirDensity = (353.05 / (t + 273)) * Exp((-0.034) * (height / (t + 273)))
        End If
    End If
End Function
