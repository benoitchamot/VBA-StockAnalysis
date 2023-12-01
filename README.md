# Stock Analysis
## Files
- CalculateYearlyChange.vbs: code for the main assignment
- CalculateMaxima.vbs: code for the first bonus
- RunAll.vbs: code for the second bonus and needs both the other files
- Clear.vbs: code to clear one or all sheets

## Data description
File: Multiple_year_stock_data.xlsx

The file contains three tabs for three different years: 2018, 2019, 2020.

Each tab contains 7 attributes (columns):
- Ticker = name of the stock
- Date in format YYYYMMDD
- Open price on the day (no currency specified)
- Highest price on the day (no currency specified)
- Lowest price on the day (no currency specified)
- Close price on the day (no currency specified)
- Volume of stocks

## Code description
All the code can be included in Multiple_year_stock_data.xlsm and has also been exported in separate VBA files for ease of access and review.

The subroutines are broken down to be run separately but can also be run all together by using the Run All (run all routines on a single tab) and Run On All Tabs buttons.

- CalculateYearlyChange (button called "Yearly Change")
- CalculateMaxima (button called "Maxima")
- Clear (button called "Clear")
- RunAll (button called "Run All")
- Clear All (button called "Clear All")

Sub CalculateYearlyChange (in file CalculateYearlyChange.vbs) corresponds to the main assignment.
Sub CalculateMaxima (in file CalculateMaxima.vbs) corresponds to the first bonus.
Sub RunAll (in file RunAll.vbs) corresponds to the second bonus and needs both the other files

Sub Clear and ClearAll are not part of the assignments but are used to conveniently reset the Excel sheet. They are provided in Clear.vbc.
