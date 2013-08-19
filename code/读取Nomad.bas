Attribute VB_Name = "读取Nomad"
' Nomad格式数据读取

Function decNomad(rs As Object)
    Dim s As Station: Set s = New Station
    
    Dim sn As String: sn = "site" & Replace(rs.Range("A3").Value, "Site Name: ", "")
    
    If sheetExist(sn) Then
        s.setSheet Sheets(sn)
    Else
        Sheets.Add After:=Sheets(Sheets.count)
        ActiveSheet.Name = sn
        s.newStation ActiveSheet
        
        decInfoNomad rs, s
    End If
    
    Sheets.Add After:=Sheets(Sheets.count)
    ActiveSheet.Name = "data" + s.id
    
    decDataNomad rs, s, ActiveSheet
    
    addStation s
End Function

Function decInfoNomad(rs As Object, s As Object)
    s.System = "Nomad"
    
    Dim nss As New Collection
    
    Dim i As Single
    For i = 1 To rs.UsedRange.Rows.count
        If InStr(1, rs.Cells(i, 1).Value, "Nomad2 Name", 1) > 0 Then
            Dim reNLS As Object: Set reNLS = CreateObject("vbscript.regexp")
            reNLS.Pattern = "Nomad2\s+Name:\s*(\d+)"
            
            Set mymatches = reNLS.Execute(rs.Cells(i, 1).Value)
            Set mymatch = mymatches(0)
            If mymatch.SubMatches.count >= 1 Then
                s.Logger.Serial = mymatch.SubMatches(0)
            End If
        ElseIf InStr(1, rs.Cells(i, 1).Value, "Site Name", 1) > 0 Then
            Dim reNSS As Object: Set reNSS = CreateObject("vbscript.regexp")
            reNSS.Pattern = "Site\s+Name:\s*(\S+)" ' \S 与 [^,#\?\/%&=]
            
            Set mymatches = reNSS.Execute(rs.Cells(i, 1).Value)
            Set mymatch = mymatches(0)
            If mymatch.SubMatches.count >= 1 Then
                s.Site.Site = mymatch.SubMatches(0)
                s.Site.Site = Replace(s.Site.Site, "#", "")
            End If
        ElseIf InStr(1, rs.Cells(i, 1).Value, "TimeStamp", 1) > 0 Then
            Dim reNSr As Object: Set reNSr = CreateObject("vbscript.regexp")
            reNSr.Pattern = "^([^\(]+)\((.+)\)(\s+@\s+(\d+)m|)[^\-]*\-\s*(\d+)\s+(min|hour)\s+(Vec\s+|)(Sampl|Averag|Max\sValu|Min\sValu|Std\sDe|Time\sOf\sMa)"
            
            Dim j As Integer
            For j = 2 To rs.UsedRange.Columns.count
                Set mymatches = reNSr.Execute(rs.Cells(i, j).Value)
                If mymatches.count > 0 Then
                    Dim ss As NomadSensor: Set ss = New NomadSensor

                    Set mymatch = mymatches(0).SubMatches
                    With ss
                        .Description = mymatch(0)
                        .Units = mymatch(1)
                        .height = mymatch(3)
                        .cat = mymatch(7)
                        .Col = j
                    End With
                    
                    Select Case ss.Units
                        Case "°"
                            ss.Units = "deg"
                        Case "°C"
                            ss.Units = "C"
                    End Select
                    
                    nss.Add ss
                End If
            Next

            s.datastart = i + 1
            Exit For
        End If
    Next i
    
    s.id = s.Site.Site
    
    getSfSN s, nss

End Function


Function getSfSN(s As Object, ns As Collection)
    Dim n As NomadSensor
    Dim k As String
    Dim ss As sSensor

    Dim i
    For i = 1 To ns.count
        Set n = ns(i)

        k = existSN(s.sensorsR, n)
        If k = "" Then
            k = CStr(s.sensorsR.count + 1)
            
            Set ss = s.newSensor
            With ss
                .height = n.height
                .Description = n.Description
                .Units = n.Units
                .channel = k
            End With
        Else
            Set ss = s.sensorsR(k)
        End If
        
        Select Case n.cat
            Case "Averag", "Sampl"
                ss.avg = n.Col
            Case "Max Valu"
                ss.max = n.Col
            Case "Min Valu"
                ss.Min = n.Col
            Case "Std De"
                ss.Sd = n.Col
        End Select
    Next
    
End Function

Function existSN(ss As Scripting.Dictionary, n As NomadSensor) As String
    Dim k
    For Each k In ss
        If isSameNomadSensor(ss(k), n) Then
            existSN = k
            Exit Function
        End If
    Next
    
    existSN = ""
End Function

Function isSameNomadSensor(s As sSensor, n As NomadSensor) As Boolean
    If s.height = n.height And s.Description = n.Description And s.Units = n.Units Then
        isSameNomadSensor = True
        Exit Function
    End If
    
    isSameNomadSensor = False
End Function


Function decDataNomad(rs As Object, s As Object, ds As Object)
    Dim maxX: maxX = rs.UsedRange.Rows.count
    
    ' Add Time
    ds.Range("A1").Value = "Date & Time Stamp"
    rangeCopy rs.Range(rs.Cells(s.datastart, 1), rs.Cells(maxX, 1)), ds.Cells(2, 1)
    ds.Range(ds.Cells(2, 1), ds.Cells(maxX - s.datastart, 1)).NumberFormatLocal = "yyyy/m/d h:mm"

    Dim i: i = 1
    
    For Each k In s.sensorsR
        Dim ss As Object: Set ss = s.sensorsR(k)
        
        Dim cAvg As Object: Set cAvg = ds.Cells(1, (i - 1) * 4 + 2)
        Dim cSD As Object: Set cSD = cAvg.Offset(0, 1)
        Dim cMax As Object: Set cMax = cAvg.Offset(0, 2)
        Dim cMin As Object: Set cMin = cAvg.Offset(0, 3)
        cAvg.Value = "CH" + CStr(i) + "Avg"
        cSD.Value = "CH" + CStr(i) + "SD"
        cMax.Value = "CH" + CStr(i) + "Max"
        cMin.Value = "CH" + CStr(i) + "Min"
        
        If ss.avg > 0 Then
            rangeCopy rs.Range(rs.Cells(s.datastart, ss.avg), rs.Cells(maxX, ss.avg)), cAvg.Offset(1, 0)
            ss.avg = cAvg.Column
        End If

        If ss.Sd > 0 Then
            rangeCopy rs.Range(rs.Cells(s.datastart, ss.Sd), rs.Cells(maxX, ss.Sd)), cSD.Offset(1, 0)
            ss.Sd = cSD.Column
        End If
        
        If ss.max > 0 Then
            rangeCopy rs.Range(rs.Cells(s.datastart, ss.max), rs.Cells(maxX, ss.max)), cMax.Offset(1, 0)
            ss.max = cMax.Column
        End If

        If ss.Min > 0 Then
            rangeCopy rs.Range(rs.Cells(s.datastart, ss.Min), rs.Cells(maxX, ss.Min)), cMin.Offset(1, 0)
            ss.Min = cMin.Column
        End If

        i = i + 1
    Next

    adjustData ds, s
End Function





