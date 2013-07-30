VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "sRation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private ch As Object
Private ca As Object
Private des As Object
Private det As Object


Public Function init(a As Object)
    Set ch = a.Offset(0, 1)
    Set ca = a.Offset(0, 2)
    Set des = a.Offset(0, 3)
    Set det = a.Offset(0, 4)

End Function

' Channel
Public Property Get Channel() As String
    Channel = ch.Value
End Property
Public Property Let Channel(ByVal v As String)
    ch.Value = v
End Property
