Attribute VB_Name = "记录"
Public LoggerIndex As Integer
Private message As Object

Function 记录初始化()
    Set message = oHome.Range("A1")

    ' 清除记录
    oRecord.Range("A2:C" & CStr(oRecord.UsedRange.Rows.Count)).Clear

    LoggerIndex = 2
End Function

Private Function Logger(err As String, str As String)

    Dim s As String
    s = CStr(LoggerIndex)
    
    oRecord.Range("A" + s).Value = Format(Now(), "mm/dd hh:MM:ss ")
    oRecord.Range("B" + s).Value = err
    oRecord.Range("C" + s).Value = str
    
    
    ' 设置显示格式
    
    Dim FontColor As String
    Dim InteriorColor As String
    
    Select Case err
        Case "「Error」"
            FontColor = "-16383844"
            InteriorColor = "13551615"
        Case "「Info」"
            FontColor = "-16752384"
            InteriorColor = "13561798"
        Case "「Warn」"
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
    Logger "「Error」", str
End Function

Function Info(str As String)
    Logger "「Info」", str
End Function

Function Warn(str As String)
    Logger "「Warn」", str
End Function


Function newMessage(str As String)
    message.Value = str
    delay 3000
    message.Value = ""
End Function

