Attribute VB_Name = "��¼"
Public LoggerIndex As Integer
Private message As Object

Function ��¼��ʼ��()
    Set message = oHome.Range("A1")

    ' �����¼
    oRecord.Range("A2:C" & CStr(oRecord.UsedRange.Rows.Count)).Clear

    LoggerIndex = 2
End Function

Private Function Logger(err As String, str As String)

    Dim s As String
    s = CStr(LoggerIndex)
    
    oRecord.Range("A" + s).Value = Format(Now(), "mm/dd hh:MM:ss ")
    oRecord.Range("B" + s).Value = err
    oRecord.Range("C" + s).Value = str
    
    
    ' ������ʾ��ʽ
    
    Dim FontColor As String
    Dim InteriorColor As String
    
    Select Case err
        Case "��Error��"
            FontColor = "-16383844"
            InteriorColor = "13551615"
        Case "��Info��"
            FontColor = "-16752384"
            InteriorColor = "13561798"
        Case "��Warn��"
            FontColor = "-16751204"
            InteriorColor = "10284031"
    End Select
    
    With oRecord.Range("A" + s + ":" + "C" + s).Font
        .Color = FontColor
        .TintAndShade = 0
    End With
    With oRecord.Range("A" + s + ":" + "C" + s).Interior
        .PatternColorIndex = xlAutomatic
        .Color = InteriorColor
        .TintAndShade = 0
    End With
    
    LoggerIndex = LoggerIndex + 1
End Function

Function Error(str As String)
    Logger "��Error��", str
End Function

Function Info(str As String)
    Logger "��Info��", str
End Function

Function Warn(str As String)
    Logger "��Warn��", str
End Function


Function newMessage(str As String)
    message.Value = str
    delay 3000
    message.Value = ""
End Function


Sub ��6()
Attribute ��6.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ��6 ��
'

'
    Range("A182:A199").Select
    Selection.NumberFormatLocal = "yyyy/m/d h:mm"
End Sub
Sub ��7()
Attribute ��7.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ��7 ��
'

'
    Range("A180:A197").Select
    
End Sub
Sub ��8()
Attribute ��8.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ��8 ��
'

'
    Columns("A:A").Select
    Selection.NumberFormatLocal = "yyyy/m/d h:mm"
    Range("C9").Select
    ActiveWindow.SmallScroll Down:=27
End Sub
Sub ��9()
Attribute ��9.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ��9 ��
'

'
    Rows("1:1").Select
    Selection.Copy
    Sheets("data-8014-1h").Select
    ActiveSheet.Paste
End Sub
