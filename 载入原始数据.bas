Attribute VB_Name = "载入原始数据"

Sub 选择文件()
    ' 导入文件对话框
    Dim fd  As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .ButtonName = "打开"
        .Title = "选择数据文件"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.add "所有", "*.*"
        .Filters.add "Nomad", "*.csv"
        .Filters.add "SDR", "*.txt"
        .InitialView = msoFileDialogViewDetails
        .Show
    End With
    
    ' 载入文件数据
    Dim i As Integer
    For i = 1 To fd.SelectedItems.Count
        Dim fp As String
        fp = fd.SelectedItems(i)
        
        ' 读取原始数据
        Dim SheetName As String
        SheetName = "raw" & CStr(i)

        Call 导入原始数据(fp, SheetName)

    Next i
End Sub

Sub 导入原始数据(path As String, SheetName As String)

    ' 先导入到临时文件，再移动到此文件中
    Workbooks.OpenText Filename:=path, _
        Origin:=936, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
        Comma:=True, Space:=False, Other:=False, TrailingMinusNumbers:=True

    Sheets(1).Select
    Sheets(1).Name = SheetName
    Sheets(1).Move After:=WB.Sheets(WB.Sheets.Count)
    Sheets(SheetName).Select

    Call 显示首页
End Sub

