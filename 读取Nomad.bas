Attribute VB_Name = "读取Nomad"

' Nomad格式数据读取

Function decInfoNomad() As String
    Dim s As New Station
    Dim ss As NomadSensor
    Dim nss As New Collection
    
    Dim reNSS As Object ' site name
    Dim reNLS As Object ' logger serial
    Dim reNSr As Object ' sensor
    Set reNSS = CreateObject("vbscript.regexp")
    Set reNLS = CreateObject("vbscript.regexp")
    Set reNSr = CreateObject("vbscript.regexp")
    reNSS.Pattern = "Site\s+Name:\s*([^,#\?\/%&=]+)"
    reNLS.Pattern = "Nomad2\s+Name:\s*(\d+)"
    reNSr.Pattern = "^([^\(]+)\((.+)\)(\s+@\s+(\d+)m|)[^\-]*\-\s*(\d+)\s+(min|hour)\s+(Vec\s+|)(Sampl|Averag|Max\sValu|Min\sValu|Std\sDe|Time\sOf\sMa)"
    
    s.System = "Nomad"
    
    Dim i As Single
    For i = 1 To ActiveSheet.UsedRange.Rows.Count
        If InStr(1, Cells(i, 1).Value, "Nomad2 Name", 1) > 0 Then
            Set s.Logger = New Logger
            
            Set mymatches = reNLS.Execute(Cells(i, 1).Value)
            Set mymatch = mymatches(0)
            If mymatch.SubMatches.Count >= 1 Then
                s.Logger.Serial = mymatch.SubMatches(0)
            End If
            
            i = i + 3
        ElseIf InStr(1, Cells(i, 1).Value, "Site Name", 1) > 0 Then
            Set s.Site = New Site
            
            Set mymatches = reNSS.Execute(Cells(i, 1).Value)
            Set mymatch = mymatches(0)
            If mymatch.SubMatches.Count >= 1 Then
                s.Site.Site = mymatch.SubMatches(0)
            End If

        ElseIf InStr(1, Cells(i, 1).Value, "TimeStamp", 1) > 0 Then
            
            Dim j As Integer
            For j = 2 To ActiveSheet.UsedRange.Columns.Count
                Set mymatches = reNSr.Execute(Cells(i, j).Value)
                If mymatches.Count > 0 Then
                    Set ss = New NomadSensor

                    Set mymatch = mymatches(0).SubMatches
                    With ss
                        .Description = mymatch(0)
                        .Units = mymatch(1)
                        .Height = mymatch(3)
                        .Cat = mymatch(7)
                        .Col = j
                    End With
                    
                    Select Case ss.Units
                        Case "°"
                            ss.Units = "deg"
                        Case "°C"
                            ss.Units = "C"
                    End Select
                    
                    nss.add ss
                End If
            Next

            s.DataStart = i + 1
            Exit For
        End If
    Next i
    
    s.id = s.Site.Site
    
    getSfSN s, nss

    addStation s

    decInfoNomad = s.id
End Function

Function getSfSN(ByRef s As Station, ns As Collection)
    Dim n As NomadSensor
    Dim k As String
    Dim ss As Sensor
    
    For Each n In ns
        k = existSN(s.SensorsR, n)
        If k = "" Then
            k = CStr(s.SensorsR.Count + 1)
            
            Set ss = New Sensor
            With ss
                .Height = n.Height
                .Description = n.Description
                .Units = n.Units
                .Channel = k
            End With

            s.SensorsR.add ss.Channel, ss
        End If
        
        Select Case n.Cat
            Case "Averag", "Sampl"
                s.SensorsR(k).Avg = n.Col
            Case "Max Valu"
                s.SensorsR(k).Max = n.Col
            Case "Min Valu"
                s.SensorsR(k).Min = n.Col
            Case "Std De"
                s.SensorsR(k).SD = n.Col
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

Function isSameNomadSensor(s As Sensor, n As NomadSensor) As Boolean
    If s.Height = n.Height And s.Description = n.Description And s.Units = n.Units Then
        isSameNomadSensor = True
        Exit Function
    End If
    
    isSameNomadSensor = False
End Function


Function decDataNomad(id As String)
    Dim rs As Object ' raw sheet
    Dim ds As Object ' new data sheet
    Set rs = ActiveSheet
    
    ' Add Title
    
    Sheets.add After:=Sheets(Sheets.Count)
    Set ds = ActiveSheet

    ds.Range("A1").Value = "Date & Time Stamp"
    
    Dim i As Integer
    i = 1
    Dim k

    Dim maxX
    maxX = ActiveSheet.UsedRange.Rows.Count

    
    For Each s In Stations(id).SensorsR
        
        ds.Cells(1, (i - 1) * 4 + 2).Value = "CH" + CStr(i) + "AVG"
        ds.Cells(1, (i - 1) * 4 + 3).Value = "CH" + CStr(i) + "SD"
        ds.Cells(1, (i - 1) * 4 + 4).Value = "CH" + CStr(i) + "MAX"
        ds.Cells(1, (i - 1) * 4 + 5).Value = "CH" + CStr(i) + "MIN"
        i = i + 1
        

        MsgBox Stations(id).SensorsR(s).Avg
        
        

    Next


    
    Dim x As String, y As String


End Function

