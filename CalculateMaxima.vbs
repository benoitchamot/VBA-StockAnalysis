Sub CalculateMaxima()
' Return the stock with the "Greatest % increase", "Greatest % decrease", and"Greatest total volume"
    Dim max_percent As Double
    Dim min_percent As Double
    Dim max_volume As LongLong

    ' Check that yearlz data are available first
    If Range("I1").Value <> "Ticker" Then
        MsgBox ("Please run Yearly Change first")
        Exit Sub
        
    ' If the data are available, calculate the maxima
    Else
        ' Calculate max and min of %
        max_percent = Application.WorksheetFunction.Max(Range("K:K").Value)
        min_percent = Application.WorksheetFunction.Min(Range("K:K").Value)
        max_volume = Application.WorksheetFunction.Max(Range("L:L").Value)
        
        Range("Q2").Value = max_percent
        Range("Q3").Value = min_percent
        Range("Q4").Value = max_volume
        
        Range("Q2:Q3").NumberFormat = "0.00%"
        Range("Q4").NumberFormat = "0"
        Range("Q4").Columns.AutoFit
        
        ' Loop through the rows of the summary table to find the min/max values
        For Row = 2 To Cells(Rows.Count, 9).End(xlUp).Row
            If Cells(Row, 11).Value = max_percent Then
                Range("P2").Value = Cells(Row, 9).Value
            End If
            
            If Cells(Row, 11).Value = min_percent Then
                Range("P3").Value = Cells(Row, 9).Value
            End If
            
            If Cells(Row, 12).Value = max_volume Then
                Range("P4").Value = Cells(Row, 9).Value
            End If
            
        Next Row
        
    End If
End Sub