VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sensor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Channel         As String
Public Cat             As String
Public Description     As String
Public Details         As String
Public SerialNumber    As String
Public Height          As Single
Public ScaleFactor     As String
Public Offset          As String
Public Units           As String
    
Private Rations()       As Ration
    
Public NotInstalled    As Boolean
    
Public Value           As String

Public Avg              As Integer
Public SD               As Integer
Public Min              As Integer
Public Max              As Integer

