VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Station"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private oS      As Object

Private sys     As Object
Private ver     As Object
Private fil     As Object

Private log     As Object
Private sit     As Object
Private senr    As Object
Private ra      As Object
Private rav     As Object
Private rad     As Object

Private dats    As Object
Private stat    As Object
Private endt    As Object

Private s10m    As Object
Private s1h     As Object
Private i_d     As Object

Public Function setSheet(ByVal s As Object)
    Set oS = s
    
    Set sys = oS.Range("A1")
    Set ver = oS.Range("B1")
    Set fil = oS.Range("E1")
    
    Set log = oS.Range("A3")
    Set sit = oS.Range("A8")
    Set senr = oS.Range("A19")
    
    Set ra = senr.Offset(oConfig.MaxSensorNum + 3, 0)
    Set rav = ra.Offset(1, 0)
    Set rad = rav.Offset(oConfig.MaxSensorNum * oConfig.StationNum + 9, 0)
    
    Set dats = oS.Range("E2")
    Set stat = oS.Range("F2")
    Set endt = oS.Range("G2")
    
    Set s10m = oS.Range("H2")
    Set s1h = oS.Range("I2")
    Set i_d = oS.Range("D2")

End Function


Public Function newStation(ByVal s As Object)
    setSheet s
    
    log.Value = "-----Logger Information-----"
    log.Offset(1, 0).Value = "Model #"
    log.Offset(2, 0).Value = "Serial #"
    log.Offset(3, 0).Value = "Hardware"
    
    sit.Value = "-----Site Information-----"
    sit.Offset(1, 0).Value = "Site #"
    sit.Offset(2, 0).Value = "Site Desc"
    sit.Offset(3, 0).Value = "Project Code"
    sit.Offset(4, 0).Value = "Project Desc"
    sit.Offset(5, 0).Value = "Site Location"
    sit.Offset(6, 0).Value = "Site Elevation"
    sit.Offset(7, 0).Value = "Latitude"
    sit.Offset(8, 0).Value = "Longitude"
    sit.Offset(9, 0).Value = "Time offset (hrs)"

    senr.Value = "-----Sensor Information-----"
    senr.Offset(1, 0).Value = "Name"
    senr.Offset(1, 1).Value = "Channel"
    senr.Offset(1, 2).Value = "Type"
    senr.Offset(1, 3).Value = "Description"
    senr.Offset(1, 4).Value = "Details"
    senr.Offset(1, 5).Value = "Serial Number"
    senr.Offset(1, 6).Value = "Height"
    senr.Offset(1, 7).Value = "Scale Factor"
    senr.Offset(1, 8).Value = "Offset"
    senr.Offset(1, 9).Value = "Units"
    senr.Offset(1, 10).Value = "NotInstalled"
    senr.Offset(1, 11).Value = "Avg"
    senr.Offset(1, 12).Value = "Sd"
    senr.Offset(1, 13).Value = "Min"
    senr.Offset(1, 14).Value = "Max"
    senr.Offset(1, 15).Value = "Scat"
    
    ra.Value = "-----Ration Information-----"
    'rav.Value = "风速相关性"
    'rad.Value = "风向相关性"
    
    oS.Range("D1").Value = "id"
    oS.Range("E1").Value = "dataStart"
    oS.Range("F1").Value = "timeStart"
    oS.Range("G1").Value = "timeEnd"
    oS.Range("H1").Value = "10m sheet"
    oS.Range("I1").Value = "1h sheet"

End Function


' system
Public Property Get System() As String
    System = sys.Value
End Property
Public Property Let System(ByVal v As String)
    sys.Value = v
End Property

' version
Public Property Get Version() As String
    Version = ver.Value
End Property
Public Property Let Version(ByVal v As String)
    ver.Value = v
End Property

' FileName
Public Property Get fileName() As String
    fileName = fil.Value
End Property
Public Property Let fileName(ByVal v As String)
    fil.Value = v
End Property

' logger
Public Property Get Logger() As Object
    Set Logger = New sLogger
    Logger.init log
End Property

' site
Public Property Get Site() As Object
    Set Site = New sSite
    Site.init sit
End Property

' sensorsR
Public Property Get SensorsR() As Scripting.Dictionary     ' Sensor
    Dim i As Integer
    Dim s As Object
    Dim a As Object

    Set SensorsR = CreateObject("Scripting.Dictionary")

    For i = 1 To oConfig.MaxSensorNum
        Set a = senr.Offset(i + 1, 0)
        If InStr(1, a.Value, "CH", 1) > 0 Then
            Set s = New sSensor
            s.init a
            SensorsR.Add s.Channel, s
        End If
    Next
    
End Property

' sensors
Public Function Sensors(c As String) As Scripting.Dictionary
    Dim i As Integer
    Dim s As Object
    Dim a As Object
    
    Set Sensors = CreateObject("Scripting.Dictionary")
    
    For i = 1 To oConfig.MaxSensorNum
        Set a = senr.Offset(i + 1, 0)
        If InStr(1, a.Value, "CH", 1) > 0 Then
            Set s = New sSensor
            s.init a
            If s.Scat = c Then
                Sensors.Add s.Channel, s
            End If
        End If
    Next
End Function


'DataStart
Public Property Get DataStart() As Integer
    DataStart = dats.Value
End Property
Public Property Let DataStart(ByVal v As Integer)
    dats.Value = v
End Property

'StartTime
Public Property Get StartTime() As String
    StartTime = stat.Value
End Property
Public Property Let StartTime(ByVal v As String)
    stat.Value = v
    stat.NumberFormatLocal = "yyyy/m/d h:mm"
End Property

' EndTime
Public Property Get EndTime() As String
    EndTime = endt.Value
End Property
Public Property Let EndTime(ByVal v As String)
    endt.Value = v
    endt.NumberFormatLocal = "yyyy/m/d h:mm"
End Property

'Sheet10m
Public Property Get Sheet10m() As String
    Sheet10m = s10m.Value
End Property
Public Property Let Sheet10m(ByVal v As String)
    s10m.Value = v
End Property

' Sheet1h
Public Property Get Sheet1h() As String
    Sheet1h = s1h.Value
End Property
Public Property Let Sheet1h(ByVal v As String)
    s1h.Value = v
End Property

' id
Public Property Get id() As String
    id = i_d.Value
End Property
Public Property Let id(ByVal v As String)
    i_d.Value = v
End Property

Public Property Get newSensor() As Object
    Dim ss As Object
    Dim i
    For i = 1 To oConfig.MaxSensorNum
        If InStr(1, senr.Offset(i + 1, 0).Value, "CH", 1) <= 0 Then
            Set ss = New sSensor
            ss.init senr.Offset(i + 1, 0)
            ss.id = "CH" & i
            Set newSensor = ss
            Exit Property
        End If
    Next i
End Property






