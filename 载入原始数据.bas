Attribute VB_Name = "����ԭʼ����"

Sub ѡ���ļ�()
    ' �����ļ��Ի���
    Dim fd  As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .ButtonName = "��"
        .Title = "ѡ�������ļ�"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.add "����", "*.*"
        .Filters.add "Nomad", "*.csv"
        .Filters.add "SDR", "*.txt"
        .InitialView = msoFileDialogViewDetails
        .Show
    End With
    
    ' �����ļ�����
    Dim i As Integer
    For i = 1 To fd.SelectedItems.Count
        Dim fp As String
        fp = fd.SelectedItems(i)
        
        ' ��ȡԭʼ����
        Dim SheetName As String
        SheetName = "raw" & CStr(i)

        Call ����ԭʼ����(fp, SheetName)

    Next i
End Sub

Sub ����ԭʼ����(path As String, SheetName As String)

    ' �ȵ��뵽��ʱ�ļ������ƶ������ļ���
    Workbooks.OpenText Filename:=path, _
        Origin:=936, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
        Comma:=True, Space:=False, Other:=False, TrailingMinusNumbers:=True

    Sheets(1).Select
    Sheets(1).Name = SheetName
    Sheets(1).Move After:=WB.Sheets(WB.Sheets.Count)
    Sheets(SheetName).Select

    Call ��ʾ��ҳ
End Sub

