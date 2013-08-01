Attribute VB_Name = "相关性"

Sub 计算相关性()
    Dim cats As New Collection
    cats.Add "wv"
    ' cats.Add "wd"
    
    Dim cat As String ' wv, wd
    Dim ca
    
    Dim k1, k2, s1 As Object, s2 As Object ' station 1, station 2
    
    Dim sss1 As Scripting.Dictionary, sss2 As Scripting.Dictionary ' sensors wv wd
    Set sss1 = CreateObject("Scripting.Dictionary")
    Set sss2 = CreateObject("Scripting.Dictionary")
    
    Dim kss1, kss2, ss1 As Object, ss2 As Object ' sensor 1, 2
    
    Dim st1 As Object, st2 As Object ' sheet 1 , sheet 2,
    Dim indexi As Integer, indexj As Integer
    Dim range1 As Object, range2 As Object
    Dim ra As Object
    
    Dim pox As Integer, poy As Integer
    Dim po As Object
    
    Dim r As Object
    
    
    
    For Each ca In cats
        cat = CStr(ca)
    
        indexi = 0
        For Each k1 In Stations
            Set s1 = Stations(k1)
            Set st1 = Sheets(s1.Sheet1h)
            Set sss1 = s1.Sensors(cat)
    
            For Each kss1 In sss1
                Set ss1 = sss1(kss1)
                
                pox = (CInt(ss1.channel) - 1) * oConfig.rax + 3
                s1.rav.Offset(pox, 0).Value = "CH" + ss1.channel

                
                indexj = 0
                For Each k2 In Stations
                    Set s2 = Stations(k2)
                    Set st2 = Sheets(s2.Sheet1h)
                    Set sss2 = s2.Sensors(cat)

                    For Each kss2 In sss2
                        Set ss2 = sss2(kss2)
                        
                        If k1 = k2 And kss1 = kss2 Then
                            'GoTo goon
                        ElseIf k1 = k2 Then
    
                            poy = (indexj + CInt(ss2.channel) - 1) * oConfig.ray + 1
                            If s1.rav.Offset(1, poy).Value = "" Then
                                s1.rav.Offset(1, poy).Value = st2.Name
                            End If
                            If s1.rav.Offset(2, poy).Value = "" Then
                                s1.rav.Offset(2, poy).Value = "CH" + ss2.channel
                            End If

                            Set range1 = st1.Range(st1.Cells(2, ss1.Avg).Address + ":" + st1.Cells(st1.UsedRange.Rows.Count, ss1.Avg).Address)
                            Set range2 = st2.Range(st2.Cells(2, ss2.Avg).Address + ":" + st2.Cells(st2.UsedRange.Rows.Count, ss2.Avg).Address)

                            Set ra = New sRation
                            Set po = s1.rav.Offset(pox, poy)
                            ra.init po, pox, poy
                            If ra.channel = "" Then
                                ra.channel = ss1.channel
                            ElseIf ra.channel <> ss1.channel Then
                                MsgBox "err1"
                            End If
                            
                            If ra.Sid = "" Then
                                ra.Sid = st2.Name
                            ElseIf ra.Sid <> st2.Name Then
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

                            st1.Select
                            Set r = Range(arrCol(CInt(ss1.channel)) + "," + arrCol(CInt(ss2.channel)))
                            'l = st1.Cells(pox, poy).Left
                            't = st1.Cells(poy, poy).Top
                            Dim chartL As Object
                            Set chartL = addLinestChart(r, st1)
                            
                            's1.os.Cells(pox, poy)
                            chartL.Parent.Cut
                            s1.os.Select
                            po.Select
                            ActiveSheet.Pictures.Paste.Select

                        Else
                            ' k1 <> k2 不同站点间的相关性
                            ' 需要取时间序列相同

                        End If

                    Next
                    
                    indexj = indexj + s2.SensorsR.Count
                Next

            Next

            indexi = indexi + s1.SensorsR.Count
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
            .Left = 100
            .Top = 12
        End With
    End With
    
    With addLinestChart.Parent
         .Height = 200 ' resize
         .Width = 250  ' resize
         .Top = 0    ' reposition
         .Left = 0   ' reposition
    End With
    
    ' st.Shapes(ActiveChart.Parent.Name).Line.Visible = msoFalse ' = =
End Function





