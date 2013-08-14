Attribute VB_Name = "系统"
Public oWB           As Object
Public oHome         As Object
Public oRecord       As Object
Public oConfig       As Object
Public oTemp         As Object
Public Stations     As New Scripting.Dictionary

Public OutputDir    As String

Public arrCol() As String

Sub 系统初始化()
    Set oWB = ActiveWorkbook
    Set oHome = oWB.Sheets("首页")
    Set oRecord = oWB.Sheets("记录")
    Set oTemp = oWB.Sheets("T")
    
    While oTemp.PivotTables.Count > 0
        oTemp.Range(oTemp.PivotTables(1).TableRange2.Address).Delete Shift:=xlUp
    Wend
    
    oTemp.Cells.ClearContents
    
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

    变量初始化
End Sub

Sub 计算初始化()
    Set Stations = Nothing
End Sub

Sub 变量初始化()
    ReDim arrCol(1 To 16) As String
    arrCol(1) = "B:B"
    arrCol(2) = "F:F"
    arrCol(3) = "J:J"
    arrCol(4) = "N:N"
    arrCol(5) = "R:R"
    arrCol(6) = "V:V"
    arrCol(7) = "Z:Z"
    arrCol(8) = "AD:AD"
    arrCol(9) = "AH:AH"
    arrCol(10) = "AL:AL"
    arrCol(11) = "AP:AP"
    arrCol(12) = "AT:AT"
    arrCol(13) = "AX:AX"
    arrCol(14) = "BB:BB"
    arrCol(15) = "BF:BF"
    arrCol(16) = "BJ:BJ"
End Sub
