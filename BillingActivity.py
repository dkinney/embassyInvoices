#!/usr/local/bin/python
import re
from datetime import datetime
from ast import literal_eval
import pandas as pd
import numpy as np
from xml.sax import ContentHandler, parse

from BillingRates import BillingRates
from openpyxl import load_workbook
from InvoiceStyles import styles
from InvoiceFormat import formatFullDetailsTab, formatActivityDataTab

baseYear = '0'
upchargeRate = 0.035

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
			df['Date'] = pd.to_datetime(df["InDate"], errors="coerce").dt.strftime("%m-%d-%Y")
			df['Date'] = df['Date'].astype(dtype='datetime64[ns]')
			
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
			
			self.dateStart = pd.to_datetime(df['Date'].min(), errors="coerce")
			self.dateEnd = pd.to_datetime(df['Date'].max(), errors="coerce")
			self.data = df

			billingRates = BillingRates(verbose=False)
			self.joinWith(billingRates)

			if verbose:
				print(f'\nRaw data from {filename}')
				print(df)
				print(f'\nDate range: {self.dateStart} to {self.dateEnd}')

	def joinWith(self, billingRates: BillingRates):
		if billingRates.data is None:
			# nothing to do
			return

		# print('Joining billing activity...')
		# print(self.data)
		# print('...with billing rates')
		# print(billingRates.data)

		joined = self.data.join(billingRates.data.set_index('EmployeeID'), on='Number', how='left', rsuffix='_rates')
		joined.drop(columns=['TcpName'], inplace=True)

		joined['Rate'] = np.where(joined['RateType'] == 'Overtime', joined['BillRateOT'], joined['BillRateReg'])
		joined['Description'] = np.where(joined['RateType'] == 'Overtime', '(Overtime)', joined['Category'])
		joined['SubCLIN'] = joined['SubCLIN'].str.replace('X', baseYear)

		# reorder the columns to be more useful
		joined = joined[['Date', 'CLIN', 'Location', 'City', 'SubCLIN', 'Category', 'Description', 'EmployeeName', 'TaskID', 'TaskName', 'Hours', 'Rate', 'HourlyRateReg', 'PostingRate', 'HazardRate']]
		self.data = joined

	def details(self, clin=None, location=None):
		df = self.data.copy()
		
		if clin is not None:
			df = df.loc[df['CLIN'] == clin]

		if location is not None:
			df = df.loc[df['Location'] == location]

		grouped = df.groupby(['Date', 'CLIN', 'Location', 'City', 'SubCLIN', 'Category', 'EmployeeName', 'TaskID', 'TaskName', 'Rate', 'HourlyRateReg', 'PostingRate', 'HazardRate'], as_index=False).agg({'Hours': 'sum'})

		pivot = grouped.pivot_table(index=[
			'Date', 'CLIN', 'Location', 'City', 'SubCLIN', 'Category', 'EmployeeName', 
			'Rate', 'HourlyRateReg', 'PostingRate', 'HazardRate'
		], columns='TaskName', values='Hours').reset_index()

		pivot.sort_values(['EmployeeName'], ascending=[True], inplace=True)

		for taskName in TaskNames.values():
			if taskName not in pivot.columns:
				pivot[taskName] = 0
			else:
				pivot[taskName] = pivot[taskName].fillna(0)

		pivot['HoursReg'] = pivot['Regular'] + pivot['LocalHoliday'] + pivot['Holiday'] + pivot['Vacation'] + pivot['Admin'] + pivot['Bereavement']
		pivot['HoursOT'] = pivot['Overtime'] + pivot['On-callOT'] + pivot['ScheduledOT'] + pivot['UnscheduledOT']
		pivot['HoursTotal'] = pivot['HoursReg'] + pivot['HoursOT']
		pivot['PostWages'] = pivot['Regular'] * pivot['HourlyRateReg'] # only use "Regular" hours for posting, not OT nor other type of regular hours
		pivot['Posting'] = pivot['PostWages'] * pivot['PostingRate']
		pivot['Hazard'] = pivot['PostWages'] * pivot['HazardRate']

		pivot = pivot[[
			'Date', 'CLIN', 'Location', 'City', 'SubCLIN', 'Category', 'EmployeeName', 
			'Regular', 'LocalHoliday', 'Holiday', 'Vacation', 'Admin', 'Bereavement', 
			'Overtime', 'On-callOT', 'ScheduledOT', 'UnscheduledOT', 
			'HoursReg', 'HoursOT', 'HoursTotal',
			'HourlyRateReg', 'PostWages', 'Posting', 'Hazard'
		]]

		pivot.sort_values(['Date', 'Location', 'City', 'SubCLIN', 'Category', 'EmployeeName'], inplace=True)

		return pivot

	def isIn(self, employeeNumberList):
		records = self.data.loc[(self.data['Number'].isin(employeeNumberList))]
		return records
	
	def startYear(self):
		return self.dateStart.year
	
	def startMonth(self):
		return self.dateStart.month
	
	def billingPeriod(self):
		startMonthName = self.dateStart.strftime('%b')
		endMonthName = self.dateEnd.strftime('%b')
		billingPeriod = f'{self.dateStart.day} {startMonthName} {self.dateStart.year} - {self.dateEnd.day} {endMonthName} {self.dateEnd.year}'
		return billingPeriod

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
		invoiceDetail = self.data

		if clin is not None:
			invoiceDetail = invoiceDetail.loc[invoiceDetail['CLIN'] == clin]

		if location is not None:
			invoiceDetail = invoiceDetail.loc[invoiceDetail['Location'] == location]

		invoiceDetail = invoiceDetail.groupby(['SubCLIN', 'Description', 'EmployeeName', 'Rate'], as_index=False).agg({'Hours': 'sum'})
		invoiceDetail['Amount'] = invoiceDetail['Hours'] * invoiceDetail['Rate']
		invoiceDetail = invoiceDetail[['SubCLIN', 'Description', 'EmployeeName', 'Hours', 'Rate', 'Amount']]
		invoiceDetail.sort_values(['SubCLIN', 'EmployeeName', 'Description'], ascending=[True, True, False], inplace=True)
		return invoiceDetail
	
	def groupedForCosts(self, clin=None, location=None):
		costDetail = self.details()

		if clin is not None:
			costDetail = costDetail.loc[costDetail['CLIN'] == clin]

		posts = costDetail.groupby(['Location', 'City'], as_index=False).agg({'Posting': 'sum'})
		posts['CLIN'] = '207'
		posts['Location'] = np.where(posts['Location'] == posts['City'], posts['Location'], posts['City'] + ', ' + posts['Location'])
		posts['City'] = 'Post'
		posts['G&A'] = posts['Posting'] * upchargeRate
		posts['Total'] = posts['Posting'] + posts['G&A']
		posts.rename(columns={'City': 'Type', 'Posting': 'Amount'}, inplace=True)
		posts = posts[['CLIN', 'Location', 'Type', 'Amount', 'G&A', 'Total']]

		hazards = costDetail.groupby(['Location', 'City'], as_index=False).agg({'Hazard': 'sum'})
		hazards['CLIN'] = '208'
		hazards['Location'] = np.where(hazards['Location'] == hazards['City'], hazards['Location'], hazards['City'] + ', ' + hazards['Location'])
		hazards['City'] = 'Hazard'
		hazards['G&A'] = hazards['Hazard'] * upchargeRate
		hazards['Total'] = hazards['Hazard'] + hazards['G&A']
		hazards.rename(columns={'City': 'Type', 'Hazard': 'Amount'}, inplace=True)
		hazards = hazards[['CLIN', 'Location', 'Type', 'Amount', 'G&A', 'Total']]

		costs = pd.concat([posts, hazards])
		costs = costs.loc[costs['Total'] > 0]
		costs.sort_values(['Location', 'Type'], ascending=[True, False], inplace=True)

		return costs

	def groupedForHoursReport(self, clin=None, location=None):
		details = self.details(clin=clin, location=location)

		details = details[[
			'City',
			'SubCLIN', 
			'EmployeeName', 
			'Regular',
			'LocalHoliday', 
			'Admin',
			'Overtime',
			'On-callOT', 
			'ScheduledOT', 
			'UnscheduledOT'
		]]

		invoiceDetail = details.groupby(['City', 'SubCLIN', 'EmployeeName'], as_index=False).agg({
			'Regular': 'sum',
			'LocalHoliday': 'sum',
			'Admin': 'sum',
			'Overtime': 'sum',
			'On-callOT': 'sum',
			'ScheduledOT': 'sum',
			'UnscheduledOT': 'sum'
		})

		invoiceDetail['Subtotal'] = invoiceDetail['Regular'] + invoiceDetail['On-callOT'] + invoiceDetail['ScheduledOT'] + invoiceDetail['UnscheduledOT'] + invoiceDetail['Overtime'] + invoiceDetail['LocalHoliday'] + invoiceDetail['Admin']
		invoiceDetail.sort_values(['City', 'SubCLIN', 'EmployeeName'], inplace=True)

		invoiceDetail.rename(columns={
			'SubCLIN': 'CLIN', 
			'EmployeeName': 'Name',
			'On-callOT': 'On-call OT',
			'ScheduledOT': 'Sched OT',
			'UnscheduledOT': 'Unschd OT',
			'LocalHoliday': 'Local Hol'
		}, inplace=True)
		
		return invoiceDetail

if __name__ == '__main__':
	import sys

	billingRates = BillingRates(verbose=False)

	activityFilename = sys.argv[1] if len(sys.argv) > 1 else None

	if activityFilename is None:
		print(f'Usage: {sys.argv[0]} <billing activity file>')
		exit()

	activity = BillingActivity(activityFilename, verbose=False)

	print(f'\nInvoice Details:')
	print(f'Date range: {activity.dateStart} to {activity.dateEnd}')
	print(activity.data)

	print(f'\nInvoice Hours Report:')
	hours = activity.groupedForHoursReport(clin='002', location='Ukraine')
	print(hours)

	exit()

	now = pd.Timestamp.now().strftime("%Y%m%d%H%M")
	outputFile = f'BillingActivity-{now}.xlsx'

	with pd.ExcelWriter(outputFile) as writer:
			activity.data.to_excel(writer, sheet_name='Details', index=False)

	# Apply formatting in place
	workbook = load_workbook(outputFile)

	for styleName in styles.keys():
			workbook.add_named_style(styles[styleName])
		
	worksheet = workbook['Details']
	formatActivityDataTab(worksheet)

	workbook.save(outputFile)
	print(f'Wrote {outputFile}')

