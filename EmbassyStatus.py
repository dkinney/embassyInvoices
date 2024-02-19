#!/usr/local/bin/python
import pandas as pd
import numpy as np
from openpyxl import load_workbook

# from BillingActivity import BillingActivity
from EmployeeTime import EmployeeTime
from BillingRates import BillingRates
from InvoiceStyles import styles
from InvoiceFormat import formatInvoiceTab, formatCostsTab, formatHoursTab, formatHoursDetailsTab, formatDetailTab, formatSummaryTab

versionString = 'v5'

Regions = {
	'001': 'Asia',
	'002': 'Europe'
}

CountryCodes = {
	'China': 'CH',
	'Hong Kong': 'HK',
	'Vietnam': 'VN',
	'Belgium': 'BE',
	'Moldova': 'MD',
	'Russia': 'RU',
	'Ukraine': 'UA',
	'NATO': 'NATO'
}

CountryApprovers = {
	'China': {
		'MES': 'Christine Rosenquist – MES Resident Manager',
		'COR': 'Michael Okamura – Senior Facility Manager – EAP COR-B'
	},
	'Hong Kong': {
		'MES': 'Christine Rosenquist – MES Resident Manager',
		'COR': 'Michael Okamura – Senior Facility Manager – EAP COR-B'
	},
	'Vietnam': {
		'MES': 'Christine Rosenquist – MES Resident Manager',
		'COR': 'Michael Okamura – Senior Facility Manager – EAP COR-B'
	},
	'Moldova': {
		'MES': 'Kevin Carroll – MES Resident Manager',
		'COR': 'Akram Elfeki - Senior Facility Manager – Moscow COR'
	},
	'Russia': {
		'MES': 'Kevin Carroll – MES Resident Manager',
		'COR': 'Akram Elfeki - Senior Facility Manager – Moscow COR'
	},
	'Ukraine': {
		'MES': 'Kevin Carroll – MES Resident Manager',
		'COR': 'Akram Elfeki - Senior Facility Manager – Moscow COR'
	},
	'NATO': {
		'MES': 'Dustin Bergee – MES Resident Manager',
		'COR': 'Robert Warner – DOS-COR-M'
	}
}

def processActivityFromFile(filename):
	billingRates = BillingRates(verbose=False)

	print(f'Processing activity from {filename}...')

	invoiceSummary = []

	if filename is None:
		print(f'Usage: {sys.argv[0]} <billing activity file>')
		return

	activity = EmployeeTime(filename, verbose=False)
	activity.joinWith(billingRates)

	# print('Time: ', activity.dateStart, ' - ', activity.dateEnd)

	startYear = activity.dateStart.strftime("%Y")
	startMonth = activity.dateStart.strftime("%m")

	locationInfo = activity.locationsByCLIN()

	# print('Location info: ', locationInfo)

	for clin in locationInfo.keys():
		region = Regions[clin]
		laborInvoiceNumber = f'STATUS-{startYear}{startMonth}'

		sheetInfo = {}
	
		##########################################################
		# Hours Report (for signatures)
		##########################################################
				
		statusReportFile = f'Status-{startYear}{startMonth}-{region}-{versionString}.xlsx'

		with pd.ExcelWriter(statusReportFile) as writer:
			for location in locationInfo[clin]:
				sheetName = f'Hours-{location}'
				# print(f'Writing hours for {location} into {sheetName}...')
				# print(data)
				data = activity.groupedForHoursReport(clin=clin, location=location)
				data.to_excel(writer, sheet_name=sheetName, startrow=0, startcol=0, header=True, index=False)

				sheetName = f'Details-{location}'
				details = activity.byDate(clin=clin, location=location)
				details.drop(columns=['Region', 'Holiday', 'Vacation', 'Bereavement', 'HoursReg', 'HoursOT'], inplace=True)
				# rename some columns for space
				details.rename(columns={
					'EmployeeName': 'Name',
					'LocalHoliday': 'Local Hol',
					'On-callOT': 'On-call OT',
					'ScheduledOT': 'Sched OT',
					'UnscheduledOT': 'Unschd OT',
					'HoursTotal': 'Subtotal',
				}, inplace=True)

				details.to_excel(writer, sheet_name=sheetName, startrow=0, startcol=0, header=True, index=False)

		hoursWorkbook = load_workbook(statusReportFile)
		for styleName in styles.keys():
			hoursWorkbook.add_named_style(styles[styleName])

		for location in locationInfo[clin]:
			worksheet = hoursWorkbook[f'Hours-{location}']
			invoiceNumber = laborInvoiceNumber + CountryCodes[location]
			
			formatHoursTab(worksheet, 
				  approvers=CountryApprovers[location], 
				  locationName=location, invoiceNumber=invoiceNumber, billingFrom=activity.billingPeriod())
			
			worksheet = hoursWorkbook[f'Details-{location}']
			formatHoursDetailsTab(worksheet, locationName=location, invoiceNumber=invoiceNumber, billingFrom=activity.billingPeriod())
		
		hoursWorkbook.save(statusReportFile)

def showResult(resultDictionary):
	result = pd.DataFrame(resultDictionary)
	# result.drop(columns=['rowsToSum'], inplace=True)
	print(result)

if __name__ == '__main__':
	import sys

	if len(sys.argv) < 2:
		print(f'Usage: {sys.argv[0]} <billing activity file>')
		sys.exit(1)

	for filename in sys.argv[1:]:
		result = processActivityFromFile(filename)	