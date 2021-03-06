Attribute VB_Name = "相关性"

Sub 计算相关性()
    Dim cats As New Collection
    cats.Add "风速"
    ' cats.Add "风向"
    
    For Each kc In cats
        Dim cat As String: cat = CStr(kc)
    
        Dim indexi As Integer: indexi = 0
        For Each k1 In Stations
            Dim s1 As Object: Set s1 = Stations(k1)
            Dim st1 As Object: Set st1 = Sheets(s1.Sheet1h)
            Dim sss1 As Scripting.Dictionary: Set sss1 = s1.sensors(cat)
    
            For Each kss1 In sss1
                Dim ss1 As Object: Set ss1 = sss1(kss1)
                
                Dim pox As Integer: pox = (CInt(ss1.channel) - 1) * oConfig.rax + 3

                Dim indexj As Integer: indexj = 0
                For Each k2 In Stations
                    Dim s2 As Object: Set s2 = Stations(k2)
                    Dim st2 As Object: Set st2 = Sheets(s2.Sheet1h)
                    Dim sss2 As Scripting.Dictionary: Set sss2 = s2.sensors(cat)

                    For Each kss2 In sss2
                        Dim ss2 As Object: Set ss2 = sss2(kss2)
                        
                        If k1 = k2 And kss1 = kss2 Then
                            'GoTo goon
                        ElseIf k1 = k2 Then
                            st1.Select
    
                            Dim poy As Integer: poy = (indexj + CInt(ss2.channel) - 1) * oConfig.ray + 1

                            Dim range1 As Object: Set range1 = Range(arrCol(CInt(ss1.channel))) ' columns(ss.avg)
                            Dim range2 As Object: Set range2 = Range(arrCol(CInt(ss2.channel))) ' ..

                            'Set range1 = st1.Range(st1.Cells(2, ss1.Avg).Address + ":" + st1.Cells(st1.UsedRange.Rows.Count, ss1.Avg).Address)
                            'Set range2 = st2.Range(st2.Cells(2, ss2.Avg).Address + ":" + st2.Cells(st2.UsedRange.Rows.Count, ss2.Avg).Address)

                            Dim ra As Object: Set ra = New sRation
                            Dim po As Object: Set po = s1.rav.Offset(pox, poy)
                            ra.init po, pox, poy
                            If ra.channel = "" Then
                                ra.channel = ss1.channel
                            ElseIf ra.channel <> ss1.channel Then
                                MsgBox "err1"
                            End If
                            
                            If ra.Sid = "" Then
                                ra.Sid = s2.os.Name
                            ElseIf ra.Sid <> s2.os.Name Then
                                MsgBox "err2"
                            End If
                            
                            If ra.Sch = "" Then
                                ra.Sch = ss2.channel
                            ElseIf ra.Sch <> ss2.channel Then
                                MsgBox "err2"
                            End If

                            ra.Rsq = Application.WorksheetFunction.Rsq(range1, range2)
                            ra.Slope = Application.WorksheetFunction.Slope(range1, range2)
                            ra.Intercept = Application.WorksheetFunction.Intercept(range1, range2)

                            Dim r As Object: Set r = Range(arrCol(CInt(ss1.channel)) + "," + arrCol(CInt(ss2.channel)))
                            ' columns(1)
                            Dim chartL As Object: Set chartL = addLinestChart(r, st1)
                            
                            chartL.Parent.Cut
                            s1.os.Select
                            po.Select
                            ActiveSheet.Pictures.Paste.Select

                        Else
                            ' k1 <> k2 不同站点间的相关性
                            ' 需要取时间序列相同进行计算

                        End If

                    Next
                    
                    indexj = indexj + s2.sensorsR.count
                Next

            Next

            indexi = indexi + s1.sensorsR.count
        Next
        
    Next

End Sub


Function addLinestChart(r As Object, st As Object) As Object
    Set addLinestChart = st.Shapes.AddChart.Chart
    
    With addLinestChart
        .ChartType = xlXYScatter
        .SetSourceData Source:=r
        .Legend.Delete
        .ChartTitle.Delete
    End With

    With addLinestChart.SeriesCollection(1)
        .MarkerStyle = 8
        .MarkerSize = 3
    End With

    With addLinestChart.SeriesCollection(1).Trendlines.Add
        .DisplayEquation = True
        .DisplayRSquared = True
        
        With .DataLabel
            .left = 100
            .top = 12
        End With
    End With
    
    With addLinestChart.Parent
         .height = 200 ' resize
         .width = 250  ' resize
         .top = 0    ' reposition
         .left = 0   ' reposition
    End With
    
    ' st.Shapes(ActiveChart.Parent.Name).Line.Visible = msoFalse ' = =
End Function





