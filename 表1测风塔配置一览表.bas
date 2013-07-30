Attribute VB_Name = "��1���������һ����"

Function showInfo(s As Station, ds As Object)
    Dim maxH As Single
    
    ds.Columns("A:A").HorizontalAlignment = xlCenter
    ds.Columns("B:B").HorizontalAlignment = xlCenter
    ds.Columns("C:C").HorizontalAlignment = xlCenter
    

    ds.Range("A1:C1").Merge
    ds.Range("A1:C1").Value = s.Site.Site + "���������һ����"
    
    ds.Range("A2").Value = "�����"
    ds.Range("B2:C2").Merge
    ds.Range("B2:C2").Value = s.Site.Site
    
    ds.Range("A3").Value = "����λ��" ' ��Ҫʱ��Ҫ����ת��
    ds.Range("B3:C3").Merge
    ds.Range("B3:C3").Value = s.Site.Latitude + "," + s.Site.Longitude

    ds.Range("A4").Value = "���θ߶�"
    ds.Range("B4:C4").Merge
    ds.Range("B4:C4").Value = CStr(s.Site.SiteElevation) + " m"
    
    ds.Range("A5").Value = "���ʱ��"
    ds.Range("B5:C5").Merge
    ds.Range("B5:C5").Value = "start��end" ' ������data��д��
    
    ds.Range("A6").Value = "����"
    ds.Range("B6:C6").Merge
    ds.Range("B6:C6").Value = "70 m" ' ��ȡheight���ֵ
    
    ' sensor ����
    ds.Range("A7").Value = "�ŵ�"
    ds.Range("B7").Value = "��װ�߶� (m)"
    ds.Range("C7").Value = "�۲���Ŀ"
    
    Dim wv As New Scripting.Dictionary
    Dim wd As New Scripting.Dictionary
    Dim p As New Scripting.Dictionary
    Dim t As New Scripting.Dictionary
    Dim h As New Scripting.Dictionary
    Dim vol As New Scripting.Dictionary

    Dim ss As Sensor
    Dim i As Integer
    i = 8
    For Each Key In s.SensorsR.Keys
        Set ss = s.SensorsR(Key)
        
        If ss.Height > maxH Then
            maxH = ss.Height
        End If
        
        Select Case ss.Units
            Case "m/s", "mph"
                wv.Add ss.Channel, ss

                addInfoSensor ds, i, ss.Channel, ss.Height, "���� (m/s)"
                i = i + 1
            Case "deg", "Degrees"
                wd.Add ss.Channel, ss
                addInfoSensor ds, i, ss.Channel, ss.Height, "���� (��)"
                i = i + 1
            Case "Volts", "v"
                vol.Add ss.Channel, ss
            Case "%RH"
                h.Add ss.Channel, ss
            Case "C", "Degrees F"
                t.Add ss.Channel, ss
                addInfoSensor ds, i, ss.Channel, ss.Height, "���� (��)"
                i = i + 1
            Case "kPa", "mb", "mB"
                p.Add ss.Channel, ss
                addInfoSensor ds, i, ss.Channel, ss.Height, "��ѹ (kpa)"
                i = i + 1
        End Select
        
    Next
    
    ds.Range("B6:C6").Value = CStr(maxH) + " m"
    
    ds.Columns("A:A").EntireColumn.AutoFit
    ds.Columns("B:B").ColumnWidth = 16
    ds.Columns("C:C").ColumnWidth = 15
    
    ' ds.Name = "info-" + s.Site.Site

End Function

Function addInfoSensor(ds As Object, i As Integer, c As String, h As Single, t As String)
    ds.Range("A" + CStr(i)).Value = "CH" + c
    ds.Range("B" + CStr(i)).Value = h
    ds.Range("C" + CStr(i)).Value = t
End Function

