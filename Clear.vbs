Sub Clear()
' Clear the summary tables
    number_of_rows = Cells(Rows.Count, 9).End(xlUp).Row
    Range("I1:L" & number_of_rows).Value = ""
    Range("I1:L" & number_of_rows).Interior.Color = xlNone
    
    Range("P2:Q4").Value = ""
End Sub

Sub ClearAll()
' Clear the summary tables in all sheets

    For I = 0 To 2

        ' Go to next sheet and call clear function
        MsgBox ("Clear 20" & (18 + I))
        Sheets("20" & (18 + I)).Select
        Call Clear
    Next I
    
    MsgBox ("All clear!")
End Sub