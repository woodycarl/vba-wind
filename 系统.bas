Attribute VB_Name = "ϵͳ"
Public WB           As Workbook
Public Home         As Object
Public RecordSheet  As Object
Public SettingSheet As Object
Public Stations     As New Scripting.Dictionary

Public OutputDir    As String

Sub ϵͳ��ʼ��()
    Set WB = ActiveWorkbook
    Set Home = WB.Sheets("��ҳ")
    Set RecordSheet = WB.Sheets("��¼")
    Set SettingSheet = WB.Sheets("����")
    
    Set Stations = Nothing
    
    OutputDir = WB.path + "\���\"
    
    ��¼��ʼ��
    ��������
    
End Sub

Function �����ʼ��()
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
