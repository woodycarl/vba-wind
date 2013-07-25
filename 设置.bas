Attribute VB_Name = "设置"
Public IndexPage    As String   ' 首页名称
Public Const MaxSensorNum = 20  ' 最大传感器数
Public CalID As Integer

Public RlostMethod  As String   ' 修订缺失数据的方法: avg | random
Public CalHeight    As Single
Public AirDensity   As Single
Public UseSetAD     As String
Public AutoRevise   As Boolean
Public Separate     As Boolean
Public MDH          As String


Sub 载入设置()
    IndexPage = SettingSheet.Range("B2").Value
    CalID = SettingSheet.Range("B3").Value

    RlostMethod = SettingSheet.Range("F2").Value
    CalHeight = SettingSheet.Range("F3").Value
    AirDensity = SettingSheet.Range("F4").Value
    UseSetAD = SettingSheet.Range("F5").Value
    AutoRevise = SettingSheet.Range("F6").Value
    Separate = SettingSheet.Range("F7").Value
    MDH = SettingSheet.Range("F8").Value
End Sub
