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

