Attribute VB_Name = "����"
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Sub ��ʾ��ҳ()
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

    Savetime = timeGetTime '���¿�ʼʱ��ʱ��
    While timeGetTime < Savetime + t 'ѭ���ȴ�
        DoEvents 'ת�ÿ���Ȩ���Ա��ò���ϵͳ�����������¼�
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

