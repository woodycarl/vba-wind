Attribute VB_Name = "工具"
'sheet
Function sheetExist(n As String) As Boolean
    Dim s
    For Each s In ActiveWorkbook.Sheets
        If s.Name = n Then
            sheetExist = True
            Exit Function
        End If
    Next
    
    sheetExist = False
End Function

Function newSheet(n As String) As Object
    Dim st: Set st = ActiveSheet

    Sheets.Add After:=Sheets(Sheets.count)
    Set newSheet = ActiveSheet
    newSheet.Name = getNewSheetName(n)
    
    st.Select
End Function

Function deleteSheet(st As Object)
    Application.DisplayAlerts = False
    st.Delete
    Application.DisplayAlerts = True
End Function

Function copySheet(st As Object, n As String) As Object
    Dim pst: Set pst = ActiveSheet
    
    st.Copy After:=Sheets(Sheets.count)
    Set copySheet = ActiveSheet
    copySheet.Name = getNewSheetName(n)
    
    pst.Select
End Function

Function getNewSheetName(n As String) As String
    If sheetExist(n) Then
        Dim nn As String
        nn = InputBox("表" + n + "已存在，输入新表名:")
        If nn = "" Then
            deleteSheet Sheets(n)
            getNewSheetName = n
        Else
            getNewSheetName = getNewSheetName(nn)
        End If
    Else
        getNewSheetName = n
    End If
End Function

' range
Function rangeCopy(r As Object, po As Object)
    r.Copy
    po.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Function

Function rangeF(dr As Object, r As Object, m As Variant)
    ' m xlMultiply, xlAdd, xlSubtract, xlDivide

    r.Copy

    dr.PasteSpecial Paste:=xlPasteAll, Operation:=m, _
                    SkipBlanks:=True, Transpose:=False
End Function
Function rangeFV(dr As Object, v As Double, m As Variant)
    Dim t As Object: Set t = newSheet("trangefv")
    Dim r As Object: Set r = t.Cells(1, 1)
    
    r.Value = v
    rangeF dr, r, m
    
    deleteSheet t
End Function

Function rangeMerge(dr As Object, Optional v As Variant = "", Optional horizontalA As Variant = xlCenter, _
        Optional verticalA As Variant = xlCenter, Optional wrapT As Boolean = True)
    With dr
        .HorizontalAlignment = horizontalA
        .VerticalAlignment = verticalA
        .WrapText = wrapT
        .Merge
        .Value = v
    End With
End Function

'pt
Function newPT(st As Object, dataRange As String, n As String) As Object
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        dataRange, Version:=xlPivotTableVersion14). _
        CreatePivotTable TableDestination:=st.Name + "!R1C1", TableName:=n, _
        DefaultVersion:=xlPivotTableVersion14
    Set newPT = st.PivotTables(n)
End Function

Function hidePTsum(pt As Object)
    With pt
        .ColumnGrand = False
        .RowGrand = False
    End With
End Function

