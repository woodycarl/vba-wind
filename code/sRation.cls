VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "sRation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private a As Object

Private id As Integer
Private r As Integer
Private s As Integer
Private i As Integer

Private delimiter As String

Private Sub Class_Initialize()
    delimiter = "," '设定分割符
    
    id = 0
    r = 1
    s = 2
    i = 3
End Sub

Public Function init(a1 As Object)
    Set a = a1
End Function

' Sid
Public Property Get Sid() As String
    Sid = a.Offset(-CInt(getV(id)) - 1, 0).Value
End Property
Public Property Let Sid(ByVal v As String)
    a.Offset(-CInt(getV(id)) - 1, 0).Value = v
End Property

' Channel
Public Property Get Channel() As Integer
    Channel = CInt(Replace(a.Offset(-CInt(getV(id)), 0).Value, "CH", ""))
End Property
Public Property Let Channel(ByVal v As Integer)
    a.Offset(-CInt(getV(id)), 0).Value = "CH" & v
End Property

' Index
Public Property Get Index() As Integer
    Index = CInt(getV(id))
End Property
Public Property Let Index(ByVal v As Integer)
    codeRation id, v
End Property

' Rsq
Public Property Get Rsq() As Double
    Rsq = CDbl(getV(r))
End Property
Public Property Let Rsq(ByVal v As Double)
    codeRation r, v
End Property

' Slope
Public Property Get Slope() As Double
    Slope = CDbl(getV(s))
End Property
Public Property Let Slope(ByVal v As Double)
    codeRation s, v
End Property

' Intercept
Public Property Get Intercept() As Double
    Intercept = CDbl(getV(i))
End Property
Public Property Let Intercept(ByVal v As Double)
    codeRation i, v
End Property

Private Function codeRation(arri As Integer, v As Variant)
    Dim arr
    arr = decRation()
    arr(arri) = v
    a.Value = Join(arr, delimiter)
End Function

Private Function decRation() As Variant
    decRation = Split(a.Value, delimiter)
End Function

Private Function getV(vi As Integer) As String
    Dim arr
    arr = decRation()
    getV = arr(vi)
End Function

