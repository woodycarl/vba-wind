Attribute VB_Name = "系统"
Public oWB          As Object
Public oHome        As Object
Public oRecord      As Object
Public oConfig      As Object
Public Stations     As New Scripting.Dictionary

Public OUTPUTDIR    As String

Sub 系统初始化()
    Set oWB = ActiveWorkbook
    Set oHome = oWB.Sheets("首页")
    Set oRecord = oWB.Sheets("记录")
    Set oConfig = New Setting
    oConfig.init oWB.Sheets("设置")
    
    Set Stations = Nothing
    
    OUTPUTDIR = oWB.path + "\输出\"
    
    记录初始化
    
    Dim st As Object, s As Station
    For Each st In Sheets
        If InStr(1, st.Name, "site", 1) > 0 Then
            Set s = New Station
            s.setSheet st
            
            Stations.Add s.id, s
        End If
    Next
End Sub


Sub 移除所有数据()
    Dim st As Object
    For Each st In Sheets
        If Not (InStr(1, st.Name, "首页", 1) > 0 Or _
                InStr(1, st.Name, "设置", 1) > 0 Or _
                InStr(1, st.Name, "记录", 1) > 0) Then
            deleteSheet st
        End If
    Next
End Sub
