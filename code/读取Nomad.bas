Attribute VB_Name = "读取Nomad"
' Nomad格式数据读取

Private Const pos As String = "A3" ' site.site单元格
Private Const pres As String = "site" ' site表前缀
Private Const pred As String = "data" ' data表前缀

Function decNomad(rs As Object)
    Dim s As Station: Set s = New Station
    
    Dim sn As String: sn = pres & Replace(rs.Range(pos).Value, "Site Name: ", "")
    
    If sheetExist(sn) Then
        s.setSheet Sheets(sn)
    Else
        Sheets.Add After:=Sheets(Sheets.Count)
        ActiveSheet.Name = sn
        s.newStation ActiveSheet
        
        decInfoNomad rs, s
    End If
    
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = pred + s.id
    
    decDataNomad rs, s, ActiveSheet
    
    addStation s
End Function

Function decInfoNomad(rs As Object, s As Object)
    s.System = "Nomad"
    
    Dim nss As New Collection
    
    Dim i As Single
    For i = 1 To rs.UsedRange.Rows.Count
        If InStr(1, rs.Cells(i, 1).Value, "Nomad2 Name", 1) > 0 Then
            Dim reNLS As Object: Set reNLS = CreateObject("vbscript.regexp")
            reNLS.Pattern = "Nomad2\s+Name:\s*(\d+)"
            
            Set mymatches = reNLS.Execute(rs.Cells(i, 1).Value)
            Set mymatch = mymatches(0)
            If mymatch.SubMatches.Count >= 1 Then
                s.Logger.Serial = mymatch.SubMatches(0)
            End If
        ElseIf InStr(1, rs.Cells(i, 1).Value, "Site Name", 1) > 0 Then
            Dim reNSS As Object: Set reNSS = CreateObject("vbscript.regexp")
            reNSS.Pattern = "Site\s+Name:\s*(\S+)" ' \S 与 [^,#\?\/%&=]
            
            Set mymatches = reNSS.Execute(rs.Cells(i, 1).Value)
            Set mymatch = mymatches(0)
            If mymatch.SubMatches.Count >= 1 Then
                s.Site.Site = mymatch.SubMatches(0)
            End If
        ElseIf InStr(1, rs.Cells(i, 1).Value, "TimeStamp", 1) > 0 Then
            Dim reNSr As Object: Set reNSr = CreateObject("vbscript.regexp")
            reNSr.Pattern = "^([^\(]+)\((.+)\)(\s+@\s+(\d+)m|)[^\-]*\-\s*(\d+)\s+(min|hour)\s+(Vec\s+|)(Sampl|Averag|Max\sValu|Min\sValu|Std\sDe|Time\sOf\sMa)"
            
            Dim j As Integer
            For j = 2 To rs.UsedRange.Columns.Count
                Set mymatches = reNSr.Execute(rs.Cells(i, j).Value)
                If mymatches.Count > 0 Then
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

            s.DataStart = i + 1
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
    For i = 1 To ns.Count
        Set n = ns(i)

        k = existSN(s.SensorsR, n)
        If k = "" Then
            k = CStr(s.SensorsR.Count + 1)
            
            Set ss = s.newSensor
            With ss
                .height = n.height
                .Description = n.Description
                .Units = n.Units
                .channel = k
            End With
        Else
            Set ss = s.SensorsR(k)
        End If
        
        Select Case n.cat
            Case "Averag", "Sampl"
                ss.avg = n.Col
            Case "Max Valu"
                ss.Max = n.Col
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
        ds.Cells(1, (i - 1) * 4 + 2).Value = "CH" + CStr(i) + "Avg"
        ds.Cells(1, (i - 1) * 4 + 3).Value = "CH" + CStr(i) + "SD"
        ds.Cells(1, (i - 1) * 4 + 4).Value = "CH" + CStr(i) + "Max"
        ds.Cells(1, (i - 1) * 4 + 5).Value = "CH" + CStr(i) + "Min"
        
        If ss.avg > 0 Then
            x = rs.Cells(s.DataStart, ss.avg).Address
            y = rs.Cells(maxX, ss.avg).Address
            
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





