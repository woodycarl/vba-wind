Attribute VB_Name = "�Ƴ�ԭʼ����"

Sub �Ƴ�raw()
    
    Application.DisplayAlerts = False
    Dim i As Integer
del:
    For i = 1 To Sheets.Count
        If InStr(1, Sheets(i).Name, "raw", 1) > 0 Then
            Sheets(i).Delete
            GoTo del
        End If
    Next i
    Application.DisplayAlerts = True
End Sub

