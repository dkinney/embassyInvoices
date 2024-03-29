#!/usr/local/bin/python
import pandas as pd
from openpyxl import load_workbook

from Config import Config
from LaborData import LaborData

from InvoiceStyles import styles
from InvoiceFormat import formatTimeByEmployee, formatTimeByDate

config = Config()

Regions = {}

# get region dictionary from config and swap keys and values
# to make it easy to look up the region name from the CLIN
for region in config.data['regions']:
	Regions[config.data['regions'][region]] = region


if __name__ == '__main__':
	import sys

	if len(sys.argv) < 2:
		print(f'Usage: {sys.argv[0]} <billing activity file>')
		sys.exit(1)

	filename = sys.argv[1]

	labor = LaborData.fromReportFile(filename)
	time = labor.time

	startYear = time.startYear()
	startMonth = time.startMonthName()
	locationInfo = time.locationsByCLIN()
	invoiceNumberValue = config.data['nextInvoiceNumber']

	for clin in locationInfo.keys():
		reportType = config.data['filenamePrefixes']['status']
		region = Regions[clin]
		pattern = f'{reportType}-{region}-{startYear}-{startMonth}'
		outputFile = f'{pattern}.xlsx'

		regionDate = time.statusByDate(clin=clin)
		regionDate.drop(columns=['PostName'], inplace=True)

		# if all hours are approved sort ascending by date, false if not
		dateAscending = regionDate['State'].eq('Approved').all()

		regionDate.sort_values(['Date', 'EmployeeName'], ascending=[dateAscending, True], inplace=True)

		regionEmployee = time.byEmployee(clin=clin)

		with pd.ExcelWriter(outputFile) as writer:
			regionEmployee.to_excel(writer, sheet_name='Employee', startrow=0, startcol=0, header=True, index=False)
			regionDate.to_excel(writer, sheet_name='Date', startrow=0, startcol=0, header=True, index=False)
			
		workbook = load_workbook(outputFile)

		for styleName in styles.keys():
			workbook.add_named_style(styles[styleName])

		worksheet = workbook['Employee']
		formatTimeByEmployee(worksheet)

		worksheet = workbook['Date']
		formatTimeByDate(worksheet)

		workbook.save(outputFile)