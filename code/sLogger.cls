VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "sLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private m As Object
Private s As Object
Private h As Object

Public Function init(a As Object)

    Set m = a.Offset(1, 1)
    Set s = a.Offset(2, 1)
    Set h = a.Offset(3, 1)
    
End Function


' model
Public Property Get Model() As Integer
    Model = m.Value
End Property
Public Property Let Model(ByVal v As Integer)
    m.Value = v
End Property

' serial
Public Property Get Serial() As String
    Serial = s.Value
End Property
Public Property Let Serial(ByVal v As String)
    s.Value = v
End Property

' hardwareRev
Public Property Get HardwareRev() As String
    HardwareRev = h.Value
End Property
Public Property Let HardwareRev(ByVal v As String)
    h.Value = v
End Property
