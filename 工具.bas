Attribute VB_Name = "工具"
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Sub 显示首页()
    x = ActiveSheet.Name
    If x <> IndexPage Then Sheets(IndexPage).Select
End Sub

Function reRep(str As String, patten As String, repStr As String) As String
    Dim re As New RegExp
    re.Pattern = patten
    reRep = re.replace(str, repStr)
End Function


Function delay(t As Double)
    Dim Savetime As Double

    Savetime = timeGetTime '记下开始时的时间
    While timeGetTime < Savetime + t '循环等待
        DoEvents '转让控制权，以便让操作系统处理其它的事件
    Wend

End Function

Function Sheet_SaveAs(s As String, f As String, n As String)
    Dim p
    p = OutputDir + f + Format(Now, " yyyymmdd hhmmssms")
    MkDir p
    
    Sheets(s).Copy

    ActiveWorkbook.SaveAs Filename:=p + "\" + n + ".xlsx", FileFormat _
        :=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close

End Function

