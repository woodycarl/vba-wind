VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Setting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Const ra_x As Integer = 15
Private Const ra_y As Integer = 6

Private calI    As Object
Private maxs    As Object
Private stan    As Object
Private usepic  As Object

Private rL      As Object
Private calH    As Object
Private airD    As Object
Private useS    As Object
Private autoR   As Object
Private s       As Object
Private mD      As Object

Public Function init(os As Object)

    Set calI = os.Range("B2")
    Set maxs = os.Range("B3")
    Set stan = os.Range("B4")
    Set usepic = os.Range("B5")
    
    Set rL = os.Range("F2")
    Set calH = os.Range("F3")
    Set airD = os.Range("F4")
    Set useS = os.Range("F5")
    Set autoR = os.Range("F6")
    Set s = os.Range("F7")
    Set mD = os.Range("F8")
End Function

' rax, ray
' 相关性图表间隔
Public Property Get rax() As Integer
    rax = ra_x
End Property
Public Property Get ray() As Integer
    ray = ra_y
End Property
'
' CalID
' 文档编号
Public Property Get CalID() As Integer
    CalID = calI.Value
End Property
Public Property Let CalID(ByVal v As Integer)
    calI.Value = v
End Property


' MaxSensorNum
' 最大传感器数
Public Property Get MaxSensorNum() As Integer
    MaxSensorNum = maxs.Value
End Property
Public Property Let MaxSensorNum(v As Integer)
    maxs.Value = v
End Property

' StationNum
' 站点数
Public Property Get StationNum() As Integer
    StationNum = stan.Value
End Property
Public Property Let StationNum(v As Integer)
    stan.Value = v
End Property

' ConvertToPic
' 生成图片
Public Property Get ConvertToPic() As Boolean
    ConvertToPic = usepic.Value
End Property
Public Property Let ConvertToPic(v As Boolean)
    usepic.Value = v
End Property

'
' RLostMethod
' 修订缺失数据的方法: avg | random
Public Property Get RLostMethod() As String
    RLostMethod = rL.Value
End Property
Public Property Let RLostMethod(ByVal v As String)
    rL.Value = v
End Property

' CalHeight
' 轮毂计算高度
Public Property Get CalHeight() As String
    CalHeight = calH.Value
End Property
Public Property Let CalHeight(ByVal v As String)
    calH.Value = v
End Property

' AirDensity
' 空气密度
Public Property Get AirDensity() As String
    AirDensity = airD.Value
End Property
Public Property Let AirDensity(ByVal v As String)
    airD.Value = v
End Property

' UseSetAD
' 强制使用空气密度设定值
Public Property Get UseSetAD() As String
    UseSetAD = useS.Value
End Property
Public Property Let UseSetAD(ByVal v As String)
    useS.Value = v
End Property

' AutoRevise
' 自动修订
Public Property Get AutoRevise() As Boolean
    AutoRevise = autoR.Value
End Property
Public Property Let AutoRevise(ByVal v As Boolean)
    autoR.Value = v
End Property

' Separate
' 多站点相关
Public Property Get Separate() As Boolean
    Separate = s.Value
End Property
Public Property Let Separate(ByVal v As Boolean)
    s.Value = v
End Property

' MDH
' 时间节点
Public Property Get MDH() As String
    MDH = mD.Value
End Property
Public Property Let MDH(ByVal v As String)
    mD.Value = v
End Property

