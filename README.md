# Monthly MIS Report Automation

Turns a messy raw sales Excel file into a clean, formatted MIS report in seconds, using Python (Pandas + openpyxl).

## Problem
Monthly sales data arrives with inconsistent region and product names, mixed date formats, amounts stored as text, duplicate orders and missing values. Cleaning it and building the report by hand takes hours every month.

## What the script does
1. Reads the raw Excel file (`data/sales_data_raw.xlsx`)
2. Cleans it: trims spaces, fixes text casing, converts text amounts to numbers, standardises mixed date formats, removes duplicate orders, drops incomplete rows
3. Builds report sheets: KPI summary, monthly revenue with growth %, and revenue by region, product and salesperson
4. Adds formatting: header styling, frozen panes, data bars, and green/red highlighting for monthly growth
5. Writes a Data Quality Log sheet showing what was fixed

## How to run
```
pip install pandas openpyxl
python generate_mis_report.py data/sales_data_raw.xlsx output/MIS_Report.xlsx
```

## Output sheets
| Sheet | Contents |
| --- | --- |
| Summary | Total revenue, orders, average order value, top region and product |
| Monthly | Revenue by month with month-on-month growth % |
| By Region / By Product / By Salesperson | Revenue, orders, average order value, share % |
| Clean Data | The cleaned dataset |
| Data Quality Log | Rows in, duplicates removed, incomplete rows removed, rows out |

## Sample data
`data/sales_data_raw.xlsx` is synthetic data created for this project (435 rows, Aug–Oct 2026) with deliberate errors, so the cleaning steps can be demonstrated. Replace it with your own file that has the same columns: Order ID, Order Date, Region, Product, Salesperson, Amount.

## Excel version (Days 2–4 of the plan)
The same report was also built manually in Excel using Power Query (cleaning), Pivot Tables and XLOOKUP. Add the workbook to this repo as `excel/MIS_Report_Manual.xlsx` and add screenshots below.

## Tools
Python, Pandas, openpyxl, Excel (Power Query, Pivot Tables, XLOOKUP, Conditional Formatting)
