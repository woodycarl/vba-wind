VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "sSite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private s   As Object
Private sde As Object
Private pc  As Object
Private pd  As Object
Private sl  As Object
Private se  As Object
Private la  As Object
Private lo  As Object
Private t   As Object

Public Function init(a As Object)

    Set s = a.Offset(1, 1)      ' "B9"
    Set sde = a.Offset(2, 1)    ' "B10"
    Set pc = a.Offset(3, 1)
    Set pd = a.Offset(4, 1)
    Set sl = a.Offset(5, 1)
    Set se = a.Offset(6, 1)
    Set la = a.Offset(7, 1)
    Set lo = a.Offset(8, 1)
    Set t = a.Offset(9, 1)
    
End Function


' site
Public Property Get Site() As String
    Site = s.Value
End Property
Public Property Let Site(ByVal v As String)
    s.Value = v
End Property

' siteDesc
Public Property Get SiteDesc() As String
    SiteDesc = sde.Value
End Property
Public Property Let SiteDesc(ByVal v As String)
    sde.Value = v
End Property

' ProjectCode
Public Property Get ProjectCode() As String
    ProjectCode = pc.Value
End Property
Public Property Let ProjectCode(ByVal v As String)
    pc.Value = v
End Property

' ProjectDesc
Public Property Get ProjectDesc() As String
    ProjectDesc = pd.Value
End Property
Public Property Let ProjectDesc(ByVal v As String)
    pd.Value = v
End Property

' SiteLocation
Public Property Get SiteLocation() As String
    SiteLocation = sl.Value
End Property
Public Property Let SiteLocation(ByVal v As String)
    sl.Value = v
End Property

' SiteElevation
Public Property Get SiteElevation() As Single
    SiteElevation = se.Value
End Property
Public Property Let SiteElevation(ByVal v As Single)
    se.Value = v
End Property

' Latitude
Public Property Get Latitude() As String
    Latitude = la.Value
End Property
Public Property Let Latitude(ByVal v As String)
    la.Value = v
End Property

' Longitude
Public Property Get Longitude() As String
    Longitude = lo.Value
End Property
Public Property Let Longitude(ByVal v As String)
    lo.Value = v
End Property

' TimeOffset
Public Property Get TimeOffset() As String
    TimeOffset = t.Value
End Property
Public Property Let TimeOffset(ByVal v As String)
    t.Value = v
End Property

