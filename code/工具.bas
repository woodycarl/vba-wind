Attribute VB_Name = "工具"

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


Function rangeCopy(r As Object, po As Object)
    r.Copy
    po.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Function


Function newSheet(n As String) As Object
    Dim st: Set st = ActiveSheet

    Sheets.Add After:=Sheets(Sheets.count)
    Set newSheet = ActiveSheet
    newSheet.Name = n
    
    st.Select
End Function

Function newPT(st As Object, dataRange As String, n As String) As Object
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        dataRange, Version:=xlPivotTableVersion14). _
        CreatePivotTable TableDestination:=st.Name + "!R1C1", TableName:=n, _
        DefaultVersion:=xlPivotTableVersion14
    Set newPT = st.PivotTables(n)
End Function
