#!/usr/local/bin/python
import sys
import pandas as pd
import calendar

from EmployeeTime import EmployeeTime
from BillingRates import BillingRates

from openpyxl import load_workbook
from InvoiceStyles import styles
from InvoiceFormat import formatTimeByDate, formatTimeByEmployee

if len(sys.argv) < 1:
    print("Usage: python3 {sys.argv[0]} <data.csv>")
    sys.exit(1)

filename = sys.argv[1]
data = pd.read_csv(filename)
data.drop(columns=['Contract ID'], inplace=True)

data['Entry Date'] = pd.to_datetime(data['Entry Date'], errors='coerce')
data.sort_values(by=['Region', 'Entry Date'], inplace=True)
data.set_index('Entry Date', inplace=True)

billingRates = BillingRates()

print('Billing Rates:')
print(billingRates.data)


for region in data['Region'].unique():
    regionData = data.loc[data['Region'] == region].copy()
    outputFile = f'{region}-Time.xlsx'
    with pd.ExcelWriter(outputFile) as writer:
        # regionData.drop(columns=['Region'], inplace=True)

        for year in regionData.index.year.unique():
            yearData = regionData.loc[regionData.index.year == year]

            for month in yearData.index.month.unique():
                monthData = yearData.loc[yearData.index.month == month].copy()
                monthData['Description'] = 'Description'
                monthData['Date'] = monthData.index
                # print(monthData)

                df = monthData[['Employee Name', 'Employee ID', 'Date', 'Description', 'Task Name', 'Duration', 'State']].copy()

                df.rename(columns={
                    'Employee Name': 'EmployeeName', 
                    'Employee ID': 'Number', 
                    'Date': 'Date',
                    'Description': 'Description',
                    'Task Name': 'TaskName', 
                    'Duration': 'Hours',
                    'State': 'State'
                }, inplace=True)
            
                # print(df)

                time = EmployeeTime()
                time.data = df
                time.joinWith(billingRates)

                timeByEmployee = time.byEmployee()
                timeByEmployee.drop(columns=['Region'], inplace=True)
                sheetName = f'{calendar.month_abbr[month]}{year}'

                timeByEmployee.to_excel(writer, sheet_name=sheetName, index=False)

    workbook = load_workbook(outputFile)

    for styleName in styles.keys():
        workbook.add_named_style(styles[styleName])

    for worksheet in workbook.worksheets:
        formatTimeByEmployee(worksheet)

    workbook.save(outputFile)