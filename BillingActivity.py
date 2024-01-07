#!/usr/local/bin/python
import re
from datetime import datetime
from ast import literal_eval
import pandas as pd
import numpy as np
from xml.sax import ContentHandler, parse

from BillingRates import BillingRates

TaskNames = {
	'3322': 'Regular',
	'3323': 'Overtime',
	'3324': 'On-callOT',
	'3325': 'ScheduledOT',
	'3326': 'UnscheduledOT',
	'3329': 'Holiday',
	'3330': 'LocalHoliday',
	'3331': 'Bereavement',
	'3332': 'Vacation',
	'3333': 'Admin'
}

RateTypes = {
	'3322': 'Regular',
	'3323': 'Overtime',
	'3324': 'Overtime',
	'3325': 'Overtime',
	'3326': 'Overtime',
	'3320': 'Regular',
	'3330': 'Regular',
	'3331': 'Regular',
	'3332': 'Regular',
	'3333': 'Regular'
}

# from https://stackoverflow.com/questions/33470130/read-excel-xml-xls-file-with-pandas
class ExcelHandler(ContentHandler):
	def __init__(self):
		self.chars = [  ]
		self.cells = [  ]
		self.rows = [  ]
		self.tables = [  ]

	def characters(self, content):
		self.chars.append(content)

	def startElement(self, name, atts):
		if name=="Cell":
			self.chars = [  ]
		elif name=="Row":
			self.cells=[  ]
		elif name=="Table":
			self.rows = [  ]

	def endElement(self, name):
		if name=="Cell":
			self.cells.append(''.join(self.chars))
		elif name=="Row":
			self.rows.append(self.cells)
		elif name=="Table":
			self.tables.append(self.rows)

class BillingActivity:
	def __init__(self, filename=None, verbose=False):
		self.data = None	# a dataframe containing the full billing information loaded from a file
		self.dateStart = datetime(1970,1,1)	# start date of the billing period loaded from a file
		self.dateEnd = datetime(3000,1,1)	# end date of the billing period loaded from a file

		if filename is not None:
			if verbose:
				print(f'Parsing billing data from {filename}')

			# we get data from Time Clock Plus (TCP) in an OpenXML format
			# since it is not .xlsx proper, we have to parse it
			excelHandler = ExcelHandler()
			parse(filename, excelHandler)
			df = pd.DataFrame(excelHandler.tables[0][4:], columns=excelHandler.tables[0][3])

			# Find label of the first row where the value 'Number' is found (within column 0)
			# This eliminates the header rows that are above the actual data
			row_label = (df.iloc[:, 0] == 'Number').idxmax()
			df = df.loc[row_label + 1:, :]

			# explicitly set the column names
			df.columns = ['Number', 'TcpName', 'Select', 'InDate', 'InTime', 'OutDate', 'OutTime', 'TaskID', 'Hours']

			# We only need the date for these hours, not the time nor the specific in/out times
			# We want the date in a datetime, not a string
			df["Date"] = pd.to_datetime(df["InDate"], errors="coerce").dt.strftime("%m-%d-%Y")
			
			# Get the TaskName and RateType from the TaskID
			df['TaskName'] = df['TaskID'].map(lambda x: TaskNames.get(x, 'Unknown'))
			df['RateType'] = df['TaskID'].map(lambda x: RateTypes.get(x, 'Unknown'))
			
			df.fillna('None', inplace=True)
			df = df[~df['Number'].isin(['Total:', 'None'])]
			df = df.replace('', np.NaN)

			df = df.dropna(axis=0, how='any', subset=['TaskID', 'Hours'])
			df['Number'].ffill(inplace=True)
			df['TcpName'].ffill(inplace=True)
			df['Hours'] = df['Hours'].astype(float)

			df.drop(columns=['Select', 'InDate', 'InTime', 'OutDate', 'OutTime'], inplace=True)
			df.sort_values(['Date', 'TcpName', 'TaskID'], ascending=[True, True, True], inplace=True)

			self.data = df

			self.dateStart = pd.to_datetime(df['Date'].min(), errors="coerce")
			self.dateEnd = pd.to_datetime(df['Date'].max(), errors="coerce")

			if verbose:
				print(f'\nRaw data from {filename}')
				print(df)
				print(f'\nDate range: {self.dateStart} to {self.dateEnd}')

	def joinWith(self, billingRates: BillingRates):
		if billingRates.data is None:
			# nothing to do
			return
		
		joined = self.data.join(billingRates.data.set_index('EmployeeID'), on='Number', how='left', rsuffix='_rates')
		joined.drop(columns=['TcpName'], inplace=True)
		joined['Rate'] = np.where(joined['RateType'] == 'Overtime', joined['BillRateOT'], joined['BillRateReg'])
		joined['Description'] = np.where(joined['RateType'] == 'Overtime', '(Overtime)', joined['Category'])

		# reorder the columns to be more useful
		joined = joined[['Date', 'CLIN', 'Location', 'City', 'SubCLIN', 'Description', 'EmployeeName', 'TaskID', 'TaskName', 'Hours', 'Rate', 'HourlyRateReg', 'PostingRate', 'HazardRate']]
		self.data = joined

	def isIn(self, employeeNumberList):
		records = self.data.loc[(self.data['Number'].isin(employeeNumberList))]
		return records
	
	def startYear(self):
		print(f'self.dateStart: {self.dateStart} type: {type(self.dateStart)}')
		return self.dateStart.year
	
	def startMonth(self):
		return self.dateStart.month
	
	def locationsByCLIN(self):
		result = {}
		for clin in self.data['CLIN'].unique():
			clinData = self.data.loc[self.data['CLIN'] == clin]

			locations = []
			for location in clinData['Location'].unique():
				locations.append(location)

			result[clin] = locations

		return result

	def groupedForInvoicing(self, clin=None, location=None):
		invoiceDetail = activity.data

		if clin is not None:
			invoiceDetail = invoiceDetail.loc[invoiceDetail['CLIN'] == clin]

		if location is not None:
			invoiceDetail = invoiceDetail.loc[invoiceDetail['Location'] == location]

		invoiceDetail = invoiceDetail.groupby(['SubCLIN', 'Description', 'EmployeeName', 'Rate'], as_index=False).agg({'Hours': 'sum'})
		invoiceDetail['Amount'] = invoiceDetail['Hours'] * invoiceDetail['Rate']
		invoiceDetail = invoiceDetail[['SubCLIN', 'Description', 'EmployeeName', 'Hours', 'Rate', 'Amount']]
		invoiceDetail.sort_values(['SubCLIN', 'EmployeeName', 'Description'], ascending=[True, True, False], inplace=True)
		return invoiceDetail

if __name__ == '__main__':
	import sys

	billingRates = BillingRates(verbose=False)

	activityFilename = sys.argv[1] if len(sys.argv) > 1 else None

	if activityFilename is None:
		print(f'Usage: {sys.argv[0]} <billing activity file>')
		exit()

	activity = BillingActivity(activityFilename, verbose=False)
	activity.joinWith(billingRates)

	print(f'\nBilling Activity:')
	print(f'Date range: {activity.dateStart} to {activity.dateEnd}')

	print(f'\nGrouped for Invoicing:')
	locationInfo = activity.locationsByCLIN()

	for clin in locationInfo.keys():
		print(f'\n\n------------------------------\nInvoices for CLIN: {clin}')

		for location in locationInfo[clin]:
			print(f'\nInvoice for: {location}')
			data = activity.groupedForInvoicing(clin=clin, location=location)
			print(data)

			summary = pd.DataFrame(columns=['SubCLIN', 'Description', 'EmployeeName', 'Hours', 'Rate', 'Amount'])
			summary.loc[0] = ['', f'Totals for {location}', '', data['Hours'].sum(), '', data['Amount'].sum()]
			print(summary)