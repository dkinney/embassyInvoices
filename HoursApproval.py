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

	for clin in locationInfo.keys():
		reportType = 'HoursApproval'
		region = Regions[clin]
		pattern = f'{reportType}-{region}-{startYear}-{startMonth}'
		outputFile = f'{pattern}.xlsx'

		print('region:', region)

		# byDate = time.dateDetails(clin=clin)
		# byEmployee = time.employeeDetails(clin=clin)
		
		with pd.ExcelWriter(outputFile) as writer:
			for country in sorted(locationInfo[clin]):
				print('country:', country)

				byEmployee = time.employeeDetails(clin=clin, location=country)
				byEmployee.to_excel(writer, sheet_name=f'Hours-{country}', startrow=0, startcol=0, header=True, index=False)

				byDate = time.dateDetails(clin=clin, location=country)
				byDate.to_excel(writer, sheet_name=f'Details-{country}', startrow=0, startcol=0, header=True, index=False)

		workbook = load_workbook(outputFile)

		for styleName in styles.keys():
			workbook.add_named_style(styles[styleName])

		for country in sorted(locationInfo[clin]):
			worksheet = workbook[f'Hours-{country}']
			# invoiceNumber = laborInvoiceNumber + CountryCodes[location]
			
			formatHoursTab(worksheet, 
				  approvers=CountryApprovers[country], 
				  locationName=country, billingFrom=time.billingPeriod())
			
			worksheet = workbook[f'Details-{country}']
			formatHoursDetailsTab(worksheet, locationName=country, billingFrom=time.billingPeriod())

		workbook.save(outputFile)