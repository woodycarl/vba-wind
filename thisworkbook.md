
Private Sub Workbook_Open()
    Call 系统初始化
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If Sh.name = "设置" Then
        Call 载入设置
    End If
End Sub