Attribute VB_Name = "¶ÁÈ¡Nomad"

Function decNomad(rs As Object)
    Dim s As Station
    Set s = New Station
    
    Dim Site As String

    Site = "site" & Replace(rs.Range("A3").Value, "Site Name: ", "")
    
    If sheetExist(Site) Then
        s.setSheet Sheets(Site)
    Else
        Sheets.Add after:=Sheets(Sheets.Count)
        ActiveSheet.Name = Site
        s.newStation Sheets(Site)
        
        decInfoNomad rs, s

        s.id = s.Site.Site
    End If
    
    Sheets.Add after:=Sheets(Sheets.Count)
    ActiveSheet.Name = "data" + s.id
    
    decDataNomad rs, s, ActiveSheet
    
    addStation s

End Function


Function decInfoNomad(rs As Object, s As Object)

    Dim ss As NomadSensor
    Dim nss As New Collection
    
    Dim reNSS As Object ' site name
    Dim reNLS As Object ' logger serial
    Dim reNSr As Object ' sensor
    Set reNSS = CreateObject("vbscript.regexp")
    Set reNLS = CreateObject("vbscript.regexp")
    Set reNSr = CreateObject("vbscript.regexp")
    reNSS.Pattern = "Site\s+Name:\s*(\S+)" ' \S Óë [^,#\?\/%&=]
    reNLS.Pattern = "Nomad2\s+Name:\s*(\d+)"
    reNSr.Pattern = "^([^\(]+)\((.+)\)(\s+@\s+(\d+)m|)[^\-]*\-\s*(\d+)\s+(min|hour)\s+(Vec\s+|)(Sampl|Averag|Max\sValu|Min\sValu|Std\sDe|Time\sOf\sMa)"
    
    s.System = "Nomad"
    
    Dim i As Single
    For i = 1 To rs.UsedRange.Rows.Count
        If InStr(1, rs.Cells(i, 1).Value, "Nomad2 Name", 1) > 0 Then
            Set mymatches = reNLS.Execute(rs.Cells(i, 1).Value)
            Set mymatch = mymatches(0)
            If mymatch.SubMatches.Count >= 1 Then
                s.Logger.Serial = mymatch.SubMatches(0)
            End If
        ElseIf InStr(1, rs.Cells(i, 1).Value, "Site Name", 1) > 0 Then
            Set mymatches = reNSS.Execute(rs.Cells(i, 1).Value)
            Set mymatch = mymatches(0)
            If mymatch.SubMatches.Count >= 1 Then
                s.Site.Site = mymatch.SubMatches(0)
            End If
        ElseIf InStr(1, rs.Cells(i, 1).Value, "TimeStamp", 1) > 0 Then
            
            Dim j As Integer
            For j = 2 To rs.UsedRange.Columns.Count
                Set mymatches = reNSr.Execute(rs.Cells(i, j).Value)
                If mymatches.Count > 0 Then
                    Set ss = New NomadSensor

                    Set mymatch = mymatches(0).SubMatches
                    With ss
                        .Description = mymatch(0)
                        .Units = mymatch(1)
                        .Height = mymatch(3)
                        .Cat = mymatch(7)
                        .col = j
                    End With
                    
                    Select Case ss.Units
                        Case "¡ã"
                            ss.Units = "deg"
                        Case "¡ãC"
                            ss.Units = "C"
                    End Select
                    
                    nss.Add ss
                End If
            Next

            s.DataStart = i + 1
            Exit For
        End If
    Next i
    
    getSfSN s, nss

End Function


Function getSfSN(s As Object, ns As Collection)
    Dim n As NomadSensor
    Dim k As String
    Dim ss As sSensor


    Dim i
    For i = 1 To ns.Count
        Set n = ns(i)

        k = existSN(s.SensorsR, n)
        If k = "" Then
            k = CStr(s.SensorsR.Count + 1)
            
            Set ss = s.newSensor
            With ss
                .Height = n.Height
                .Description = n.Description
                .Units = n.Units
                .Channel = k
            End With
        Else
            Set ss = s.SensorsR(k)
        End If
        
        Select Case n.Cat
            Case "Averag", "Sampl"
                ss.Avg = n.col
            Case "Max Valu"
                ss.Max = n.col
            Case "Min Valu"
                ss.Min = n.col
            Case "Std De"
                ss.Sd = n.col
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
    If s.Height = n.Height And s.Description = n.Description And s.Units = n.Units Then
        isSameNomadSensor = True
        Exit Function
    End If
    
    isSameNomadSensor = False
End Function


Function decDataNomad(rs As Object, s As Object, ds As Object)
    Dim maxX
    maxX = rs.UsedRange.Rows.Count
    Dim x, y
    
    ' Add Title
    ds.Range("A1").Value = "Date & Time Stamp"
    
    x = rs.Cells(s.DataStart, 1).Address
    y = rs.Cells(maxX, 1).Address
    
    rs.Range(x + ":" + y).Copy
    ds.Cells(2, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    
    
    Dim i As Integer
    i = 1
    Dim ss As Object
    Dim k

    For Each k In s.SensorsR
        Set ss = s.SensorsR(k)
        ds.Cells(1, (i - 1) * 4 + 2).Value = "CH" + CStr(i) + "AVG"
        ds.Cells(1, (i - 1) * 4 + 3).Value = "CH" + CStr(i) + "SD"
        ds.Cells(1, (i - 1) * 4 + 4).Value = "CH" + CStr(i) + "MAX"
        ds.Cells(1, (i - 1) * 4 + 5).Value = "CH" + CStr(i) + "MIN"
        
        If ss.Avg > 0 Then
            x = rs.Cells(s.DataStart, ss.Avg).Address
            y = rs.Cells(maxX, ss.Avg).Address
            
            rs.Range(x + ":" + y).Copy
            ds.Cells(2, (i - 1) * 4 + 2).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
        End If
        
        
        If ss.Sd > 0 Then
            x = rs.Cells(s.DataStart, ss.Sd).Address
            y = rs.Cells(maxX, ss.Sd).Address
            
            rs.Range(x + ":" + y).Copy
            ds.Cells(2, (i - 1) * 4 + 3).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
        End If
        
        If ss.Max > 0 Then
            x = rs.Cells(s.DataStart, ss.Max).Address
            y = rs.Cells(maxX, ss.Max).Address
            
            rs.Range(x + ":" + y).Copy
            ds.Cells(2, (i - 1) * 4 + 4).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
        End If

        If ss.Min > 0 Then
            x = rs.Cells(s.DataStart, ss.Min).Address
            y = rs.Cells(maxX, ss.Min).Address
            
            rs.Range(x + ":" + y).Copy
            ds.Cells(2, (i - 1) * 4 + 5).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
        End If

        i = i + 1

    Next

    adjustData ds, s
    
End Function





