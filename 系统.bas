Attribute VB_Name = "ϵͳ"
Public oWB           As Object
Public oHome         As Object
Public oRecord       As Object
Public oConfig       As Object
Public Stations     As New Scripting.Dictionary

Public OutputDir    As String

Sub ϵͳ��ʼ��()
    Set oWB = ActiveWorkbook
    Set oHome = oWB.Sheets("��ҳ")
    Set oRecord = oWB.Sheets("��¼")
    
    Set oConfig = New Setting
    oConfig.init oWB.Sheets("����")
    
    Set Stations = Nothing
    
    OutputDir = oWB.path + "\���\"
    
    '��¼��ʼ��
    ��¼��ʼ��
    
    Dim sh As Object, s As Station
    For Each sh In Sheets
        If InStr(1, sh.Name, "site", 1) > 0 Then
            Set s = New Station
            s.setSheet sh
            
            Stations.Add s.id, s
        End If
    Next
    
End Sub

Sub �����ʼ��()
    Set Stations = Nothing
End Sub
