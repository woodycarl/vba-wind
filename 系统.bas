Attribute VB_Name = "系统"
Public WB           As Workbook
Public Home         As Object
Public RecordSheet  As Object
Public SettingSheet As Object
Public Stations     As New Scripting.Dictionary

Public OutputDir    As String

Sub 系统初始化()
    Set WB = ActiveWorkbook
    Set Home = WB.Sheets("首页")
    Set RecordSheet = WB.Sheets("记录")
    Set SettingSheet = WB.Sheets("设置")
    
    Set Stations = Nothing
    
    OutputDir = WB.path + "\输出\"
    
    记录初始化
    载入设置
    
End Sub

Function 计算初始化()
    Set Stations = Nothing
End Function

Function sheetExist(n As String) As Boolean
    Dim s
    For Each s In WB.Sheets
        If s.Name = n Then
            sheetExist = True
            Exit Function
        End If
    Next
    
    sheetExist = False
End Function
