VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "sSensor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private i_d As Object

Private ch As Object
Private ca As Object
Private des As Object
Private det As Object
Private ser As Object
Private sen As Object
Private h As Object
Private scf As Object
Private off As Object
Private uni As Object
' ration
Private noi As Object

Private av As Object
Private s As Object
Private mi As Object
Private ma As Object

Private sca As Object

Public Function init(a As Object)
    Set i_d = a
    Set ch = a.Offset(0, 1)
    Set ca = a.Offset(0, 2)
    Set des = a.Offset(0, 3)
    Set det = a.Offset(0, 4)
    Set sen = a.Offset(0, 5)
    Set h = a.Offset(0, 6)
    Set scf = a.Offset(0, 7)
    Set off = a.Offset(0, 8)
    Set uni = a.Offset(0, 9)
    Set noi = a.Offset(0, 10)
    Set av = a.Offset(0, 11)
    Set s = a.Offset(0, 12)
    Set mi = a.Offset(0, 13)
    Set ma = a.Offset(0, 14)
    Set sca = a.Offset(0, 15)
End Function

' id
Public Property Get id() As String
    id = i_d.Value
End Property
Public Property Let id(ByVal v As String)
    i_d.Value = v
End Property

' Channel
Public Property Get channel() As String
    channel = ch.Value
End Property
Public Property Let channel(ByVal v As String)
    ch.Value = v
End Property

' Cat
Public Property Get cat() As String
    cat = ca.Value
End Property
Public Property Let cat(ByVal v As String)
    ca.Value = v
End Property

' Description
Public Property Get Description() As String
    Description = des.Value
End Property
Public Property Let Description(ByVal v As String)
    des.Value = v
End Property

' Details
Public Property Get Details() As String
    Details = det.Value
End Property
Public Property Let Details(ByVal v As String)
    det.Value = v
End Property

' SerialNumber
Public Property Get SerialNumber() As String
    SerialNumber = sen.Value
End Property
Public Property Let SerialNumber(ByVal v As String)
    sen.Value = v
End Property

' Height
Public Property Get height() As Single
    height = h.Value
End Property
Public Property Let height(ByVal v As Single)
    h.Value = v
End Property

' ScaleFactor
Public Property Get ScaleFactor() As String
    ScaleFactor = scf.Value
End Property
Public Property Let ScaleFactor(ByVal v As String)
    scf.Value = v
End Property

' Offset
Public Property Get Offset() As String
    Offset = off.Value
End Property
Public Property Let Offset(ByVal v As String)
    off.Value = v
End Property

' Units
Public Property Get Units() As String
    Units = uni.Value
End Property
Public Property Let Units(ByVal v As String)
    uni.Value = v
End Property

' NotInstalled
Public Property Get NotInstalled() As Boolean
    NotInstalled = noi.Value
End Property
Public Property Let NotInstalled(ByVal v As Boolean)
    noi.Value = v
End Property

' Avg
Public Property Get avg() As Integer
    avg = av.Value
End Property
Public Property Let avg(ByVal v As Integer)
    av.Value = v
End Property

' SD
Public Property Get Sd() As Integer
    Sd = s.Value
End Property
Public Property Let Sd(ByVal v As Integer)
    s.Value = v
End Property

' Min
Public Property Get Min() As Integer
    Min = mi.Value
End Property
Public Property Let Min(ByVal v As Integer)
    mi.Value = v
End Property

' Max
Public Property Get max() As Integer
    max = ma.Value
End Property
Public Property Let max(ByVal v As Integer)
    ma.Value = v
End Property

' Scat wv wd t p
Public Property Get Scat() As String
    Scat = sca.Value
End Property
Public Property Let Scat(ByVal v As String)
    sca.Value = v
End Property


'
' Ration
'
Public Property Get Rations() As Scripting.Dictionary
    Set Rations = CreateObject("Scripting.Dictionary")
    
    Dim i
    For i = 1 To oConfig.MaxSensorNum * oConfig.StationNum
        
    
    Next
    
End Property




