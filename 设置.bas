Attribute VB_Name = "����"
Public IndexPage    As String   ' ��ҳ����
Public Const MaxSensorNum = 20  ' ��󴫸�����
Public CalID As Integer

Public RlostMethod  As String   ' �޶�ȱʧ���ݵķ���: avg | random
Public CalHeight    As Single
Public AirDensity   As Single
Public UseSetAD     As String
Public AutoRevise   As Boolean
Public Separate     As Boolean
Public MDH          As String


Sub ��������()
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
