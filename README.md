# VBA-challenge
## Description
VBA challenge for Monash University Bootcamp Module 2

## Files
- Multiple_year_stock_data.xlsx: the original data file, unchanged)
- Multiple_year_stock_data.xlsm: a copy of the original file with macro enabled and the new code
- CalculateYearlyChange.vb: code for the main assignment.
- CalculateMaxima.vbs: code for the first bonus.
- RunAll.vbs: code for the second bonus and needs both the other files

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
File: Multiple_year_stock_data.xlsm

All the code is included in this Excel document and has also been exported in separate VBA files for ease of access and review.

The subroutines are broken down to be run separately but can also be run all together by using the Run All (run all routines on a single tab) and Run On All Tabs buttons.

- CalculateYearlyChange (button called "Yearly Change")
- CalculateMaxima (button called "Maxima")
- Clear (button called "Clear")
- RunAll (button called "Run All")
- Clear All (button called "Clear All")

Sub CalculateYearlyChange (in file CalculateYearlyChange.vbs) corresponds to the main assignment.
Sub CalculateMaxima (in file CalculateMaxima.vbs) corresponds to the first bonus.
Sub RunAll (in file RunAll.vbs) corresponds to the second bonus and needs both the other files

Sub Clear and ClearAll are not part of the assignments but are used to conveniently reset the Excel sheet. They are not provided in separate VBS files.
