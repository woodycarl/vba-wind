Attribute VB_Name = "test"

    
Sub test()
    Dim dic As Scripting.Dictionary
    Set dic = CreateObject("Scripting.Dictionary")
    dic.Add "Avg", "风速 (m/s)"
    dic.Add "WP", "风功率密度 (W/m2)"
    
    MsgBox VarType(dic("Avg"))
End Sub
