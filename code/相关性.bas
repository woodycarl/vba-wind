Attribute VB_Name = "相关性"

Sub 计算相关性()
    Dim cats As New Collection
    cats.Add "wv"
    cats.Add "wd"
    
    Dim k, s As Object
    Dim cat
    Dim sss As Scripting.Dictionary
    For Each k In Stations
        Set s = Stations(k)
        
        For Each cat In cats
            sss = s.Sensors(cat)
            
            calLinest s, sss
        Next
        
    Next

End Sub

Function calLinest(s As Object, sss As Scripting.Dictionary)
    Dim ch1, ch2
    Dim ss1, ss2
    Dim k1, k2
    
    For Each ss1 In sss
        For Each ss2 In sss
        
            If ss1.Channel < ss2.Channel Then
                
            
        
            End If
        Next
    Next
End Function
