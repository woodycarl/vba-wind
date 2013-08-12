Attribute VB_Name = "输出威布尔曲线"


Private Function showWeibulls()

    Dim wvs As Object: Set wvs = s.Sensors("wv")
    Dim av: av = wvs.Items
    
    For i = 0 To wvs.Count - 1
        Dim ss As Object: Set ss = av(i)

        showWeibull rst, dst, s, po, pt, ss, ssd
        
        Set po = po.Offset(25, 0)
    Next
End Function

Private Function showWeibull()
    Dim maxX As Integer: maxX = rst.UsedRange.Rows.Count
    Dim range1 As Object: Set range1 = rst.Range(rst.Cells(2, ss.avg).Address + ":" + rst.Cells(maxX, ss.avg).Address)

    Dim sr As Double
    sr = Sqr(Application.WorksheetFunction.DevSq(range1) / (maxX - 1))
    Dim avg As Double
    avg = Application.WorksheetFunction.Average(range1)
    
    k = (sv / avg) ^ (-1.086)
    c = avg / gamma(1 + 1 / k)
    
    'Gamma_Dist
    
End Function


Private Function gamma(v As Double) As Double
    gamma = Exp(Application.WorksheetFunction.GammaLn(v))
End Function

Private Function weibull(k As Double, c As Double, v As Double) As Double
    weibull = Application.WorksheetFunction.weibull(v, k, c, False)
End Function
