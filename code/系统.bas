Attribute VB_Name = "系统"
Public oWB           As Object
Public oHome         As Object
Public oRecord       As Object
Public oConfig       As Object
Public Stations     As New Scripting.Dictionary

Public OutputDir    As String

Sub 系统初始化()
    Set oWB = ActiveWorkbook
    Set oHome = oWB.Sheets("首页")
    Set oRecord = oWB.Sheets("记录")
    
    Set oConfig = New Setting
    oConfig.init oWB.Sheets("设置")
    
    Set Stations = Nothing
    
    OutputDir = oWB.path + "\输出\"
    
    '记录初始化
    记录初始化
    
    Dim sh As Object, s As Station
    For Each sh In Sheets
        If InStr(1, sh.Name, "site", 1) > 0 Then
            Set s = New Station
            s.setSheet sh
            
            Stations.Add s.id, s
        End If
    Next
    
End Sub

Sub 计算初始化()
    Set Stations = Nothing
End Sub
