Attribute VB_Name = "test"

Sub test()

    
    MsgBox Range("A10").Top
    MsgBox ActiveWindow.PointsToScreenPixelsY(0)
End Sub
