Attribute VB_Name = "����ԭʼ�ļ�"

Sub ѡ���ļ�()
    ' open choose file dialog
    Dim fd  As FileDialog
    Dim fp As String
    Dim sn As String ' sheet name
    Dim i As Integer ' raw index
    i = 1
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .ButtonName = "��"
        .Title = "Choose Data File"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "All", "*.*"
        .Filters.Add "Nomad", "*.csv"
        .Filters.Add "SDR", "*.txt"
        .InitialView = msoFileDialogViewDetails
        .Show
        
        For Each f In .SelectedItems
            fp = CStr(f)
            
            delExistRaw fp      ' ɾ���ѵ������ͬ�ļ�
            sn = newRawName(i)  ' ������raw sheet��
            
            ����ԭʼ�ļ� fp, sn
            
            i = i + 1
        Next
    End With
    
End Sub

Function ����ԭʼ�ļ�(path As String, SheetName As String)
    Dim fs As Object
    Set fs = ActiveSheet
    ' open in tmp file, move to oWB
    Workbooks.OpenText FileName:=path, _
        Origin:=936, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
        Comma:=True, Space:=False, Other:=False, TrailingMinusNumbers:=True

    Sheets(1).Select
    Sheets(1).Name = SheetName
    Sheets(1).Move after:=oWB.Sheets(oWB.Sheets.Count)
    Sheets(SheetName).Select
    Range("E1").Value = "FileName"
    Range("F1").Value = path
    
    fs.Select
End Function

Function delExistRaw(path As String)
    Application.DisplayAlerts = False
    
    Dim s
del:
    For Each s In Sheets
        If InStr(1, s.Name, "raw", 1) > 0 Then
            If s.Range("F1") = path Then
                s.Delete
                GoTo del
            End If
        End If
    Next
    
    Application.DisplayAlerts = True
End Function

Function newRawName(i As Integer) As String
    If sheetExist("raw" & i) Then
        newRawName = newRawName(i + 1)
    Else
        newRawName = "raw" & i
    End If
End Function

Sub �Ƴ�ԭʼ����()
    Application.DisplayAlerts = False
    
    Dim s
del:
    For Each s In Sheets
        If InStr(1, s.Name, "raw", 1) > 0 Then
            s.Delete
            GoTo del
        End If
    Next
    
    Application.DisplayAlerts = True
End Sub

