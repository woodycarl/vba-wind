Attribute VB_Name = "载入原始文件"
' 打开文件选择对话框
' 选择文件后自动导入,并生成以raw开头的表

Private Const cfp As String = "F1" ' 用于存放文件名的单元格编号
Private Const cfpn As String = "E1" ' 提示
Private Const preRaw As String = "raw" ' 表名前缀

Sub 选择文件()
    系统初始化
    
    Dim fd  As Object
    Dim fp As String
    Dim sn As String ' sheet name
    Dim i As Integer ' raw index
    i = 1
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .ButtonName = "打开"
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
            
            delExistRaw fp      ' 删除已导入的相同文件
            sn = newRawName(i)  ' 生成新raw sheet名
            
            loadData fp, sn
            
            i = i + 1
        Next
    End With
    
End Sub

Private Function loadData(path As String, SheetName As String)
    Dim fs As Object
    Set fs = ActiveSheet
    ' open in tmp file, move to oWB
    Workbooks.OpenText fileName:=path, _
        Origin:=936, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
        Comma:=True, Space:=False, Other:=False, TrailingMinusNumbers:=True

    Sheets(1).Select
    Sheets(1).Name = SheetName
    Sheets(1).Move after:=oWB.Sheets(oWB.Sheets.Count)
    
    Sheets(SheetName).Select
    Range(cfpn).Value = "FileName"
    Range(cfp).Value = path
    
    fs.Select
End Function

Private Function delExistRaw(path As String)
    Application.DisplayAlerts = False
    
    Dim s
del:
    For Each s In Sheets
        If InStr(1, s.Name, preRaw, 1) > 0 Then
            If s.Range(cfp) = path Then
                s.Delete
                GoTo del
            End If
        End If
    Next
    
    Application.DisplayAlerts = True
End Function

Private Function newRawName(i As Integer) As String
    If sheetExist(preRaw & i) Then
        newRawName = newRawName(i + 1)
    Else
        newRawName = preRaw & i
    End If
End Function

Sub 移除原始数据()
    Application.DisplayAlerts = False
    
    Dim st
del:
    For Each s In Sheets
        If InStr(1, st.Name, preRaw, 1) > 0 Then
            st.Delete
            GoTo del
        End If
    Next
    
    Application.DisplayAlerts = True
End Sub

