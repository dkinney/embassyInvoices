#!/usr/local/bin/python
import pandas as pd
from openpyxl import load_workbook

from Config import Config
from LaborData import LaborData

from InvoiceStyles import styles
from InvoiceFormat import formatHoursTab, formatHoursDetailsTab

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

		reportType = 'HoursStatus'
		pattern = f'{reportType}-{region}-{startYear}-{startMonth}'
		outputFile = f'{pattern}.xlsx'

		sheetInfo = {}

		with pd.ExcelWriter(outputFile) as writer:
			for locationName in sorted(invoiceData.locationDetails.keys()):
				locationData = invoiceData.locationDetails[locationName]

				sheetName = f'Hours-{locationName}'
				for item in locationData.hoursSummary:
					item.to_excel(writer, sheet_name=sheetName, startrow=0, startcol=0, index=False, header=True)

				sheetName = f'Details-{locationName}'
				for item in locationData.hoursDetail:
					item.to_excel(writer, sheet_name=sheetName, startrow=0, startcol=0, index=False, header=True)
		
		workbook = load_workbook(outputFile)

		for styleName in styles.keys():
			workbook.add_named_style(styles[styleName])

		for locationName in sorted(invoiceData.locationDetails.keys()):
			worksheet = workbook[f'Hours-{locationName}']
			# invoiceNumber = laborInvoiceNumber + CountryCodes[location]
			
			formatHoursTab(worksheet, 
				  approvers=CountryApprovers[locationName], 
				  locationName=locationName, billingFrom=time.billingPeriod())
			
			worksheet = workbook[f'Details-{locationName}']
			formatHoursDetailsTab(worksheet, locationName=locationName, billingFrom=time.billingPeriod())

		workbook.save(outputFile)