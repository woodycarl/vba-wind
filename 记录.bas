Attribute VB_Name = "¼ÇÂ¼"
Public LoggerIndex As Integer
Private message As Object

Function ¼ÇÂ¼³õÊ¼»¯()
    Set message = oHome.Range("A1")

    ' Çå³ý¼ÇÂ¼
    oRecord.Range("A2:C" & CStr(oRecord.UsedRange.Rows.Count)).Clear

    LoggerIndex = 2
End Function

Private Function Logger(err As String, str As String)

    Dim s As String
    s = CStr(LoggerIndex)
    
    oRecord.Range("A" + s).Value = Format(Now(), "mm/dd hh:MM:ss ")
    oRecord.Range("B" + s).Value = err
    oRecord.Range("C" + s).Value = str
    
    
    ' ÉèÖÃÏÔÊ¾¸ñÊ½
    
    Dim FontColor As String
    Dim InteriorColor As String
    
    Select Case err
        Case "¡¸Error¡¹"
            FontColor = "-16383844"
            InteriorColor = "13551615"
        Case "¡¸Info¡¹"
            FontColor = "-16752384"
            InteriorColor = "13561798"
        Case "¡¸Warn¡¹"
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
    Logger "¡¸Error¡¹", str
End Function

Function Info(str As String)
    Logger "¡¸Info¡¹", str
End Function

Function Warn(str As String)
    Logger "¡¸Warn¡¹", str
End Function


Function newMessage(str As String)
    message.Value = str
    delay 3000
    message.Value = ""
End Function


Sub ºê6()
Attribute ºê6.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ºê6 ºê
'

'
    Range("A182:A199").Select
    Selection.NumberFormatLocal = "yyyy/m/d h:mm"
End Sub
Sub ºê7()
Attribute ºê7.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ºê7 ºê
'

'
    Range("A180:A197").Select
    
End Sub
Sub ºê8()
Attribute ºê8.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ºê8 ºê
'

'
    Columns("A:A").Select
    Selection.NumberFormatLocal = "yyyy/m/d h:mm"
    Range("C9").Select
    ActiveWindow.SmallScroll Down:=27
End Sub
Sub ºê9()
Attribute ºê9.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ºê9 ºê
'

'
    Rows("1:1").Select
    Selection.Copy
    Sheets("data-8014-1h").Select
    ActiveSheet.Paste
End Sub
