Attribute VB_Name = "��ȡ"

Public reISH    As Object   ' info sensor height
Public re1      As Object   ' �ж�ʱ������1:
Public re2      As Object   ' �ж�ʱ������2:
Public re3      As Object

Sub ��ȡ����()
    initRead
    
    Dim st As Object
    
    For Each st In Sheets
        If InStr(1, st.Name, "raw", 1) > 0 Then
            decRaw st
        End If
    Next
    
    Dim k, s As Object
    For Each k In Stations
        Set s = Stations(k)
        adjustR s
    Next
    
End Sub

Sub ����1h()

    Dim k, st As Object
    For Each k In Stations
        Set st = Stations(k)
        
        If st.Sheet1h = "" Then
            If st.Sheet10m = "" Then
                MsgBox st.id + ": no data!"
                Exit Sub
            End If
            Info st.id + " no 1h data! gen from 10m data."
            
            Sheets.Add after:=Sheets(Sheets.Count)
            ActiveSheet.Name = "data" + st.id + "1h"
            
            genD1fD2 Sheets(st.Sheet10m), ActiveSheet
            
            st.Sheet1h = "data" + st.id + "1h"
        End If
    Next
End Sub

' �趨����
Function initRead()
    Set reISH = CreateObject("vbscript.regexp")
    reISH.Pattern = "^([\d\.]+)\s*(m|ft)"

    Set re1 = CreateObject("vbscript.regexp")
    re1.Pattern = "(\d{4})[\/|-](\d{1,2})[\/|-](\d{1,2})(\s\w+|)\s(\d{1,2}):(\d{1,2})(:\d{1,2}|)"

    Set re2 = CreateObject("vbscript.regexp")
    re2.Pattern = "(\d{1,2})[\/|-](\d{1,2})[\/|-](\d{4})\s(\d{1,2}):(\d{1,2})(:\d{1,2}|)"
    
    Set re3 = CreateObject("vbscript.regexp")
    re3.Pattern = "^""?(\d{4})[\-|\/](\d{1,2})[\-|\/](\d{1,2})""?$"
End Function

Function decRaw(s As Object)
    Dim v
    v = s.Range("A1").Value
    
    If InStr(1, v, "SDR", 1) > 0 Then

        decSDR s
        'decDataSDR decInfoSDR
    ElseIf InStr(1, v, "Multi-Track Export -", 1) > 0 Then
        decNomad s
        'decDataNomad decInfoNomad
    End If
End Function

Function addStation(s As Station)

    If Stations.Count >= 1 Then
        If Stations.Exists(s.id) Then
            Info "վ���Ѵ���: " + CStr(s.id)
            Exit Function
        End If
    End If

    Info "����վ��: " + s.id
    Stations.Add s.id, s
        
End Function


Function decDate(str As String) As String
    Set mymatches = re1.Execute(str)
    If mymatches.Count >= 1 Then
        Set mymatch = mymatches(0)
        If mymatch.SubMatches.Count >= 6 Then
            decDate = newDate(mymatch.SubMatches(0), mymatch.SubMatches(1), mymatch.SubMatches(2), mymatch.SubMatches(4), mymatch.SubMatches(5))
            Exit Function
        End If
    End If
    
    Set mymatches = re2.Execute(str)
    If mymatches.Count >= 1 Then
        Set mymatch = mymatches(0)
        If mymatch.SubMatches.Count >= 6 Then
            decDate = newDate(mymatch.SubMatches(2), mymatch.SubMatches(0), mymatch.SubMatches(1), mymatch.SubMatches(3), mymatch.SubMatches(4))
            Exit Function
        End If
    End If
    
    Set mymatches = re3.Execute(str)
    If mymatches.Count >= 1 Then
        Set mymatch = mymatches(0)
        If mymatch.SubMatches.Count >= 3 Then
            decDate = newDate(mymatch.SubMatches(0), mymatch.SubMatches(1), mymatch.SubMatches(2), 0, 0)
            Exit Function
        End If
    End If
    
    MsgBox "time format err: " + str
    'Error "ʱ���ʽ����"
    
End Function

Function newDate(y As Integer, mo As Integer, d As Integer, h As Integer, Min As Integer) As String
    newDate = CStr(y) + "/" + CStr(mo) + "/" + CStr(d) + " " + CStr(h) + ":" + CStr(Min)
End Function

Function adjustData(ds As Object, s As Object)
    ' �������ڸ�ʽ
    
    Dim i
    For i = 2 To ds.UsedRange.Rows.Count
        ds.Cells(i, 1).Value = decDate(ds.Cells(i, 1).Value)
    Next i
    
    ds.Columns("A:A").NumberFormatLocal = "yyyy/m/d h:mm"
    
    ' �ж���10���ӻ���60��������
    
    Dim maxX, maxY
    maxX = ds.UsedRange.Rows.Count
    maxY = ds.UsedRange.Columns.Count
    
    Dim x As String, y As String
    x = ds.Cells(2, maxY + 1).Address
    y = ds.Cells(maxX, maxY + 1).Address
    
    ds.Cells(2, maxY + 1).Formula = "=MINUTE(A2)"
    ds.Cells(2, maxY + 1).AutoFill Destination:=ds.Range(x + ":" + y)

    ds.Cells(1, maxY + 1).Formula = "=AVERAGE(" + CStr(x) + ":" + CStr(y) + ")"
    
    Dim n As String
    If ds.Cells(1, maxY + 1).Value > 1 Then
        n = "data" + s.id + "10m"
        s.Sheet10m = n
    Else
        n = "data" + s.id + "1h"
        s.Sheet1h = n
    End If
    If sheetExist(n) Then
        Sheets(n).Delete
    End If
    ds.Name = n
    
    ds.Range(ds.Cells(1, maxY + 1).Address + ":" + ds.Cells(maxX, maxY + 1).Address).Clear
    
    ' ��ʼ����ʱ��
    
    ds.Cells(2, maxY + 1).Formula = "=min('" + ds.Name + "'!A:A)"
    ds.Cells(2, maxY + 1).NumberFormatLocal = "yyyy/m/d h:mm"
    s.StartTime = ds.Cells(2, maxY + 1).Value
    ds.Cells(2, maxY + 1).Formula = "=max('" + ds.Name + "'!A:A)"
    ds.Cells(2, maxY + 1).NumberFormatLocal = "yyyy/m/d h:mm"
    s.EndTime = ds.Cells(2, maxY + 1).Value
    
    ds.Cells(2, maxY + 1).Clear
End Function


Function genD1fD2(d2 As Object, d1 As Object)
    Dim maxX, maxY
    maxY = d2.UsedRange.Columns.Count
    maxX = d2.UsedRange.Rows.Count

    d2.Range(d2.Cells(1, 1), d2.Cells(1, maxY)).Copy
    d1.Paste
    
    Dim d2s, d2e, d1i, val
    
    Dim i, j, k
    Dim sum
    For i = 2 To maxX
        If i = 2 Then
            d2s = 2
            d2e = 2
            d1i = 2
            val = Format(d2.Cells(i, 1).Value, "yyyy/mm/dd hh")
        End If
    
        If Format(d2.Cells(i, 1).Value, "yyyy/mm/dd hh") = val Then
            d2e = d2e + 1
        Else
            d1.Cells(d1i, 1).Value = val + ":00"
            
            
            For j = 2 To maxY

                If d2.Cells(d2s, j).Value <> "" Then
                    
                    sum = 0
                    For k = d2s To d2e
                        sum = sum + d2.Cells(k, j).Value
                    Next k
                    d1.Cells(d1i, j) = sum / (d2e - d2s + 1)
                    'd1.Cells(d1i, j).Formula = "=average(" + d2.Name + "!" + d2.Cells(d2s, j).Address + ":" + d2.Cells(d2e, j).Address + ")"
                End If
            Next j
            
            d2s = i
            d2e = i
            d1i = d1i + 1
            val = Format(d2.Cells(i, 1).Value, "yyyy/mm/dd hh")
        End If
        
    Next i
    
    d1.Range(Cells(2, 2), Cells(d1.UsedRange.Rows.Count, d1.UsedRange.Columns.Count)).NumberFormatLocal = "0.00_);[��ɫ](0.00)"

End Function


Function adjustR(s As Object)

    Dim k, ss As sSensor
    For Each k In s.SensorsR
        Set ss = s.SensorsR(k)
        
        Select Case ss.Units
        Case "mph":
            adjustRTimes s.Sheet10m, ss, 1.6 / 3.6
            adjustRTimes s.Sheet1h, ss, 1.6 / 3.6
        Case "Degrees F", "F":
            adjustRF s.Sheet10m, ss
            adjustRF s.Sheet1h, ss
        Case "mb", "mB", "MB":
            adjustRTimes s.Sheet10m, ss, 0.1
            adjustRTimes s.Sheet1h, ss, 0.1
        End Select
    Next

End Function

Function adjustRTimes(sn As String, ss As Object, t As Double)
    Dim i, ds As Object
    
    If sn <> "" And sheetExist(sn) Then
        Set ds = Sheets(sn)
        For i = 2 To ds.Rows.Count
            If ss.Avg > 0 Then
                ds.Cells(i, ss.Avg).Value = ds.Cells(i, ss.Avg).Value * t
            End If
            
            If ss.Max > 0 Then
                ds.Cells(i, ss.Max).Value = ds.Cells(i, ss.Max).Value * t
            End If
            
            If ss.Min > 0 Then
                ds.Cells(i, ss.Min).Value = ds.Cells(i, ss.Min).Value * t
            End If
        Next i
    End If
End Function

Function adjustRF(sn As String, ss As Object)
    Dim i, ds As Object
    
    If sn <> "" And sheetExist(sn) Then
        Set ds = Sheets(sn)
        For i = 2 To ds.Rows.Count
            If ss.Avg > 0 Then
                ds.Cells(i, ss.Avg).Value = (ds.Cells(i, ss.Avg).Value - 32) / 1.8
            End If
            
            If ss.Max > 0 Then
                ds.Cells(i, ss.Max).Value = (ds.Cells(i, ss.Max).Value - 32) / 1.8
            End If
            
            If ss.Min > 0 Then
                ds.Cells(i, ss.Min).Value = (ds.Cells(i, ss.Min).Value - 32) / 1.8
            End If
        Next i
    End If
End Function



