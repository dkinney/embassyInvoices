#!/usr/local/bin/python
import pandas as pd
from openpyxl import load_workbook
from itertools import repeat

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

	if len(sys.argv) < 4:
		print(f'Usage: {sys.argv[0]} <countryName> <invoiceNumber> <billing activity file>')
		sys.exit(1)

	countryName = sys.argv[1]
	invoiceNumberInput = sys.argv[2]	# this is the number part, although it can have an R if a revision of a revision
	filename = sys.argv[3]

	# the invoiceNumberValue is the number part of the invoice number (before any R suffix)
	try:
		invoiceNumberValue = int(invoiceNumberInput.split('R')[0])
	except ValueError:
		print(f'Error: {invoiceNumberInput} is not a valid invoice number')
		print(f'It must be a number with zero or more R suffixes (for revisions)')
		sys.exit(1)

	# the number of R's is the revision count
	revisionCount = invoiceNumberInput.count('R') + 1

	invoiceNumberSuffix = 'R' * revisionCount
	invoiceNumber = f'SD-{invoiceNumberValue:04d}{invoiceNumberSuffix}'

	print(f'This is revision: {revisionCount}')
	print(f'Invoice Number: {invoiceNumber}')

	labor = LaborData.fromReportFile(filename)
	time = labor.time

	startYear = time.startYear()
	startMonth = time.startMonthName()
	locationInfo = time.locationsByCLIN()

	# reverse the dictionary so we can look up the CLIN from the region name
	countryToCLIN = {}
	for clin in locationInfo.keys():
		for country in locationInfo[clin]:
			countryToCLIN[country] = clin

	if countryName not in countryToCLIN.keys():
		print(f'Error: "{countryName}" is not a valid country name')
		print(f'Available countries are: {", ".join(countryToCLIN.keys())}')
		exit(-1)

	clin = countryToCLIN[countryName]
	invoiceData = labor.invoiceData[clin]

	locationData = invoiceData.locationDetails[countryName]
	prefix = config.data['filenamePrefixes']['revisedLabor']
	pattern = f'{prefix}-{countryName}-{startYear}-{startMonth}'
	outputFile = f'{pattern}.xlsx'

	sheetInfo = {}

	with pd.ExcelWriter(outputFile) as writer:
		sheetName = f'Labor-{countryName}'
		summaryStartRow = 22
		rowsToSum = []

		for item in locationData.laborDetails:
			item.to_excel(writer, sheet_name=sheetName, startrow=summaryStartRow, startcol=0, header=False)
			rowsToSum.append((summaryStartRow + 1, summaryStartRow + len(item)))
			summaryStartRow += len(item) + 2

		summary = summaryDataframe(f'Totals for {countryName}', locationData.laborHours, locationData.laborAmount)
		summary.to_excel(writer, sheet_name=sheetName, startrow=summaryStartRow, startcol=0, header=False)

		# DO NOT increment since we are only handling a single invoice invoiceNumberValue += 1
		billingPeriod = time.billingPeriod()

		invoiceDetail = {
			'description': f'{time.dateStart.strftime("%B")} {startYear}',
			'filename': outputFile,
			'type': 'Labor',
			'invoiceNumber': invoiceNumber,
			'isRevision': 'True',
			'taskOrder': f'Labor-{countryName}',
			'billingPeriod': billingPeriod,
			'invoiceAmount': summary['Amount'].sum(),
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