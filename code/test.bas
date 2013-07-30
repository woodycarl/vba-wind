Attribute VB_Name = "test"

Sub test()

    
    
    Dim ra As Object
    Set ra = New sRation
    ra.init Sheets(1).Range("B19")
    
    ra.Index = 2
    
    MsgBox ra.Channel
    MsgBox ra.Sid

    
End Sub
