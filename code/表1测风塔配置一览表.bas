Attribute VB_Name = "表1测风塔配置一览表"

Function showInfo(s As Object, dst As Object)
    Dim po As Object: Set po = s.Pc

    twoCol dst, po.Offset(1, 0), "测风塔", s.id
    twoCol dst, po.Offset(2, 0), "地理位置", s.Site.Latitude + "," + s.Site.Longitude
    twoCol dst, po.Offset(3, 0), "海拔高度", CStr(s.Site.SiteElevation) + " m"
    twoCol dst, po.Offset(4, 0), "测风时段", Format(s.StartTime, "yyyy/m/d") + "～" + Format(s.EndTime, "yyyy/m/d")
    Dim maxH As Double: maxH = Application.WorksheetFunction.max(s.os.Range("G21:G" & 20 + s.sensorsR.count))
    twoCol dst, po.Offset(5, 0), "塔高", CStr(maxH) + " m"

    po.Offset(6, 0).Value = "信道"
    po.Offset(6, 1).Value = "安装高度 (m)"
    po.Offset(6, 2).Value = "观测项目"

    Dim i As Integer: i = 7
    Dim sensors: Set sensors = s.sensorsR
    For Each k In sensors
        Dim ss As Object: Set ss = sensors(k)

        Select Case ss.Scat
            Case "风速"
                addInfoSensor po.Offset(i, 0), ss, "风速 (m/s)"
                i = i + 1
            Case "风向"
                addInfoSensor po.Offset(i, 0), ss, "风向 (度)"
                i = i + 1
            Case "气温"
                addInfoSensor po.Offset(i, 0), ss, "气温 (℃)"
                i = i + 1
            Case "气压"
                addInfoSensor po.Offset(i, 0), ss, "气压 (kpa)"
                i = i + 1
        End Select
    Next
    
    rangeMerge dr:=dst.Range(po, po.Offset(0, 2)), v:=s.id + "测风塔配置一览表"
    
    s.Pc = s.Pc.Offset(i + 2, 0)
End Function

Private Function twoCol(dst As Object, po As Object, t As Variant, v As Variant)
    po.Value = t
    rangeMerge dr:=dst.Range(po.Offset(0, 1), po.Offset(0, 2)), v:=v
End Function

Private Function addInfoSensor(po As Object, ss As Object, t As String)
    po.Value = "CH" + ss.channel
    po.Offset(0, 1).Value = ss.height
    po.Offset(0, 2).Value = t
End Function



