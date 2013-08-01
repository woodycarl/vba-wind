VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "sRation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private r As Object
Private s As Object
Private i As Object

Private c As Object
Private si As Object
Private sc As Object


Public Function init(a As Object, pox As Integer, poy As Integer)
    Set r = a
    Set s = a.Offset(1, 0)
    Set i = a.Offset(2, 0)
    
    Set c = a.Offset(0, -poy)
    Set si = a.Offset(-pox + 1, 0)
    Set sc = a.Offset(-pox + 2, 0)
    
End Function

' Sid
Public Property Get Sid() As String
    Sid = si.Value
End Property
Public Property Let Sid(ByVal v As String)
    si.Value = v
End Property

' Sch
Public Property Get Sch() As String
    Sch = Replace(sc.Value, "CH", "")
End Property
Public Property Let Sch(ByVal v As String)
    sc.Value = "CH" & v
End Property

' Channel
Public Property Get channel() As String
    channel = Replace(c.Value, "CH", "")
End Property
Public Property Let channel(ByVal v As String)
    c.Value = "CH" & v
End Property


' Rsq
Public Property Get Rsq() As Double
    Rsq = r.Value
End Property
Public Property Let Rsq(ByVal v As Double)
    r.Value = v
End Property

' Slope
Public Property Get Slope() As Double
    Slope = s.Value
End Property
Public Property Let Slope(ByVal v As Double)
    s.Value = v
End Property

' Intercept
Public Property Get Intercept() As Double
    Intercept = i.Value
End Property
Public Property Let Intercept(ByVal v As Double)
    i.Value = v
End Property

