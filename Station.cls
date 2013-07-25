VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Station"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public id           As String

Public System       As String
Public Version      As String
    
Public SensorsR     As New Scripting.Dictionary ' Sensor
Public Logger       As Logger
Public Site         As Site

Public sensors      As New Scripting.Dictionary ' sensor
    
Public DataStart    As Integer

Public StartTime    As Date
Public EndTime      As Date

