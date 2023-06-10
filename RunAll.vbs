Sub RunAll()
' Run all scripts on all sheets

    For I = 0 To 2

        ' Go to next sheet and call yearly change and maxima functions
        MsgBox ("Run scripts on 20" & (18 + I))
        Sheets("20" & (18 + I)).Select
        
        Call CalculateYearlyChange
        Call CalculateMaxima
    Next I
End Sub