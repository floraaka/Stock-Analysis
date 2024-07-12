# Stock Data Analysis with VBA in Excel
This VBA script is designed to analyze stock data across multiple quarters in Excel, calculating metrics such as quarterly change, percentage change, and total volume for each stock. It also applies conditional formatting to highlight cells with the greatest percentage increase, decrease, and volume.

![alt text](https://github.com/floraaka/Stock-Analysis/blob/main/Screenshot 2024-07-10 123611.png?raw=true)

## Features
### Processing Stock Data:

Calculates quarterly change ($), percentage change (%), and total stock volume based on opening and closing prices.
Iterates through each worksheet (quarter) in the workbook to perform calculations.
Highlighting Greatest Metrics:

### Applies conditional formatting to highlight cells:
Green for the greatest percentage increase.
Red for the greatest percentage decrease.
Optionally, applies different formatting for the greatest total volume.
Usage
### Preparing Your Workbook:

Ensure your Excel workbook (alphabetical_testing.xlsx) contains data organized in separate sheets representing different quarters.
### Running the Script:

Open Excel and the workbook (alphabetical_testing.xlsx).
Press Alt + F8 to open the macro dialog.
Select ProcessStockData from the list and click Run to execute the script.
The script will loop through each quarter, calculate metrics, and apply conditional formatting.
### Result Verification:

Check each worksheet after running the script to ensure:
Columns for quarterly change ($), percentage change (%), and total stock volume are populated.
Cells with the greatest increase, decrease, and volume are highlighted as expected.
Customization
### Column Adjustment:
Modify col = 8 in the HighlightGreatest subroutine to match the column number where your percentage change data is located.
Adjust RGB color codes (RGB(0, 255, 0) for green, RGB(255, 0, 0) for red) in HighlightGreatest for custom formatting.
Files
StockAnalysisVBA.bas: Contains the complete VBA script for processing stock data and applying conditional formatting.
README.md: This file providing instructions and information about using the script.
Notes
Ensure macros are enabled in Excel for the script to run.
Test the script on smaller datasets first (alphabetical_testing.xlsx) to ensure functionality before applying to larger datasets.
