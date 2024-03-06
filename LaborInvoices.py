#!/usr/local/bin/python
import pandas as pd
from openpyxl import load_workbook

from Config import Config
from LaborData import LaborData

from InvoiceStyles import styles
from InvoiceFormat import formatInvoiceTab

config = Config()
Regions = {}

versionString = f'v{config.data["version"]}'
CountryApprovers = config.data['approvers']

# get region dictionary from config and swap keys and values
# to make it easy to look up the region name from the CLIN
for region in config.data['regions']:
	Regions[config.data['regions'][region]] = region

# construct a summary dataframe for writing into the file
def summaryDataframe(description:str, hours:float, amount:float) -> pd.DataFrame:
	data = {
		'blank1': '',
		'Description': description,
		'blank2': '',
		'Hours': hours,
		'blank3': '',
		'Amount': amount
	}

	return pd.DataFrame(data, index=[0])

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
		region = Regions[clin]
		invoiceData = labor.invoiceData[clin]

		prefix = config.data['filenamePrefixes']['laborInvoices']
		pattern = f'{prefix}-{region}-{startYear}-{startMonth}'
		outputFile = f'{pattern}.xlsx'

		sheetInfo = {}

		with pd.ExcelWriter(outputFile) as writer:
			for locationName in sorted(invoiceData.locationDetails.keys()):
				if locationName not in locationInfo[clin] or locationName == 'Unknown':
					print(f'\n----------\nWarning: {locationName} is not in the locationInfo dictionary\n----------\n')
					continue

				locationData = invoiceData.locationDetails[locationName]
				sheetName = f'Labor-{locationName}'
				summaryStartRow = 22
				rowsToSum = []

				for item in locationData.laborDetails:
					item.to_excel(writer, sheet_name=sheetName, startrow=summaryStartRow, startcol=0, header=False)
					rowsToSum.append((summaryStartRow + 1, summaryStartRow + len(item)))
					summaryStartRow += len(item) + 2

				summary = summaryDataframe(f'Totals for {locationName}', locationData.laborHours, locationData.laborAmount)
				summary.to_excel(writer, sheet_name=sheetName, startrow=summaryStartRow, startcol=0, header=False)

				invoiceNumber = f'SD-{invoiceNumberValue:04d}'
				invoiceNumberValue += 1
				billingPeriod = time.billingPeriod()

				invoiceDetail = {
					'description': f'{time.dateStart.strftime("%B")} {startYear}',
					'region': locationName,
					'filename': outputFile,
					'type': 'Labor',
					'invoiceNumber': invoiceNumber,
					'taskOrder': f'Labor-{locationName}',
					'billingPeriod': billingPeriod,
					'invoiceAmount': 0,
					'rowsToSum': rowsToSum
				}

				sheetInfo[sheetName] = invoiceDetail
		
		workbook = load_workbook(outputFile)

		for styleName in styles.keys():
			workbook.add_named_style(styles[styleName])

		for key in sheetInfo.keys():
			worksheet = workbook[key]
			info = sheetInfo[key]
			formatInvoiceTab(worksheet, info)

		workbook.save(outputFile)