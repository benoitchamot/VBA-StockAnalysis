Sub CalculateYearlyChange()
' Loop through all the stocks for one year and output
' - the ticker symbol
' - the early change in the price
' - the percent chnage
' - the total stock volume

    ' Column Index (for easy reference in code
    ' ----------------------------------------
    Dim COLUMN_TICKER As Integer
    COLUMN_TICKER = 1
    Dim COLUMN_DATE As Integer
    COLUMN_DATE = 2
    Dim COLUMN_OPEN As Integer
    COLUMN_OPEN = 3
    Dim COLUMN_HIGH  As Integer
    COLUMN_HIGH = 4
    Dim COLUMN_LOW As Integer
    COLUMN_LOW = 5
    Dim COLUMN_CLOSE  As Integer
    COLUMN_CLOSE = 6
    Dim COLUMN_VOL As Integer
    COLUMN_VOL = 7
    ' ----------------------------------------

    ' Variable definition
    ' ----------------------------------------
    Dim number_of_rows As Long          ' total number of rows in the sheet
    Dim ticker_counter As Integer       ' counter of unique ticker
    Dim total_stock_volume As LongLong  ' total number of one stock exchanged
    Dim year_value_open As Double       ' value of stock at start of year (open)
    Dim year_value_close As Double      ' value of stock at end of year (close)
    ' ----------------------------------------

    ' Create the headers for the summary table
    ' ----------------------------------------
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    Range("I1").Columns.AutoFit
    Range("J1").Columns.AutoFit
    Range("K1").Columns.AutoFit
    Range("L1").Columns.AutoFit
    ' ----------------------------------------

    ' Get the number of rows on the sheet
    ' ----------------------------------------
    number_of_rows = Cells(Rows.Count, 1).End(xlUp).Row
    ' ----------------------------------------
    
    ' Initialise variables
    ticker_counter = 1                  ' start counting tickers at 1
    total_stock_volume = 0              ' set stock volume to 0 (loop will override)
    year_value_open = Range("C2").Value ' assign open value of first stock
    year_value_close = 0                ' set close value to 0 (loop will override)

    ' Loop through
    For Row = 2 To number_of_rows
        total_stock_volume = total_stock_volume + Cells(Row, COLUMN_VOL).Value
        year_value_close = Cells(Row, COLUMN_CLOSE).Value
        
        If Cells(Row, COLUMN_TICKER).Value <> Cells(Row + 1, COLUMN_TICKER).Value Then
            ' Add ticker to summary table
            Range("I" & (ticker_counter + 1)).Value = Cells(Row, COLUMN_TICKER).Value
            
            ' Calculate yearly change
            Range("J" & (ticker_counter + 1)).Value = year_value_close - year_value_open
            Range("K" & (ticker_counter + 1)).Value = (year_value_close - year_value_open) / year_value_open
            
            ' Apply formatting
            ' ----------------------------------------
            If Range("J" & (ticker_counter + 1)).Value < 0 Then
                ' Colour in red if change is less than 0
                Range("J" & (ticker_counter + 1)).Interior.Color = RGB(255, 0, 0)
            ElseIf Range("J" & (ticker_counter + 1)).Value > 0 Then
                ' Colour in green if change is greater than 0
                Range("J" & (ticker_counter + 1)).Interior.Color = RGB(0, 255, 0)
            End If
            
            ' Change format of relative change to percent
            Range("K" & (ticker_counter + 1)).NumberFormat = "0.00%"
            ' ----------------------------------------
            
            ' Add stock volume to summary table
            Range("L" & (ticker_counter + 1)).Value = total_stock_volume
            
            ' Reset counters
            ticker_counter = ticker_counter + 1
            total_stock_volume = 0
            year_value_open = Cells(Row + 1, COLUMN_OPEN).Value
        End If

    Next Row
    
    MsgBox ("Done")

End Sub