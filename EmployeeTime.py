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
from InvoiceFormat import formatTimeByDate, formatTimeByEmployee

baseYear = '0'
upchargeRate = 0.35

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

TaskMap = {
	# '3526': '3333',
	# '3527': '3330',
	# '3311': '3322',
	# '3314': '3325',
	# '3317': '3326',
	# '3524': '3326',
    'Scheduled Overtime': 'ScheduledOT',
	'Scheduled - Overtime': 'ScheduledOT',
	'Unscheduled Overtime': 'UnscheduledOT',
	'Unscheduled/ Emergency OT': 'UnscheduledOT',
	'On Call- Overtime': 'On-callOT',
	'Local Holiday': 'LocalHoliday'
}

RateTypes = {
	'Regular': 'Regular',
	'Overtime': 'Overtime',
	'On-callOT': 'Overtime',
	'ScheduledOT': 'Overtime',
	'UnscheduledOT': 'Overtime',
	'Holiday': 'Non-Billable',
	'LocalHoliday': 'Regular',
	'Bereavement': 'Non-Billable',
	'Vacation': 'Non-Billable',
	'Admin': 'Regular'
}

def formatName(text):
	# print(f'formatName({text})')
	tokens = text.split(' ')

	if len(tokens) < 2:
			# not a name
			return text

	lastName = tokens[0]
	firstName = tokens[1]

	middleInitial = ''

	# hacktastic - eat extra spaces
	if len(tokens) > 3:
			middleInitial = tokens[3]
	elif len(tokens) > 2:
			middleInitial = tokens[2]

	if middleInitial != '':
		return f'{lastName}, {firstName} {middleInitial}'

	return f'{lastName}, {firstName} {middleInitial}'

def cleanupTask(text):
	if text in TaskMap:
		return TaskMap[text]
	
	return text
	
class EmployeeTime:
	def __init__(self, filename=None, verbose=False):
		self.data = None	# a dataframe containing the full billing information loaded from a file
		self.asia = None
		self.europe = None
		self.dateStart = datetime(1970,1,1)	# start date of the billing period loaded from a file
		self.dateEnd = datetime(3000,1,1)	# end date of the billing period loaded from a file

		if filename is not None:
			if verbose:
				print(f'Parsing activity data from {filename}')
			
			converters = {
				'EmployeeName': str,
				'Date': datetime,
				'Description': str,
				'TaskName': str,
				'Hours': float
			}
		
			df = pd.read_excel(filename, header=5, converters=converters)
			df.columns = ['EmployeeName', 'Date', 'Description', 'TaskName', 'Hours', 'State']

			# fill down the missing EmployeeName values
			df['EmployeeName'] = df['EmployeeName'].fillna(method='ffill')

			# remove rows where the Hours are missing
			df = df.dropna(subset=['Hours'])

			# we only care about the rows that start with our contract number in the Description
			df = df.loc[df['Description'].str.startswith('19AQMM23C0047')]
			df['Region'] = np.where(df['Description'].str.contains('Asia'), 'Asia', 'Europe')

			# strip whitespace from all string columns
			df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

			# clean up values 
			df['EmployeeName'] = df['EmployeeName'].apply(formatName)
			df['Hours'] = pd.to_numeric(df['Hours'], errors="coerce")
			df['TaskName'] = df['TaskName'].apply(cleanupTask)
			df['Date'] = pd.to_datetime(df['Date'], errors="coerce")
			df['RateType'] = df['TaskName'].map(lambda x: RateTypes.get(x, 'Unknown'))

			df.sort_values(['Date'], ascending=[True], inplace=True)
			self.dateStart = df['Date'].min()
			self.dateEnd = df['Date'].max()

			if verbose:
				print(f'Loaded {len(df)} entries from {filename} from {self.dateStart} to {self.dateEnd}')
				print(df.info())
				print(df)

			self.data = df
			
	def joinWith(self, billingRates: BillingRates):
		if billingRates.data is None:
			# nothing to do
			return

		# Intacct data does not have an EmployeeID, so we need to join on the EmployeeName
		# joined = self.data.join(billingRates.data.set_index('EmployeeID'), on='Number', how='left', rsuffix='_rates')
		joined = self.data.join(billingRates.data.set_index('EmployeeName'), on='EmployeeName', how='left', rsuffix='_rates')

		joined['Rate'] = np.where(joined['RateType'] == 'Overtime', joined['BillRateOT'], joined['BillRateReg'])
		joined['Description'] = np.where(joined['RateType'] == 'Overtime', '(Overtime)', joined['Category'])
		joined['SubCLIN'] = joined['SubCLIN'].str.replace('X', baseYear)

		# reorder the columns to be more useful
		joined = joined[['Date', 'CLIN', 'Region', 'Location', 'City', 'SubCLIN', 'Category', 'Description', 'EmployeeName', 'TaskName', 'Hours', 'State', 'Rate', 'HourlyRateReg', 'PostingRate', 'HazardRate']]
		self.data = joined

	def details(self, clin=None, location=None):
		df = self.data.copy()
		
		if clin is not None:
			df = df.loc[df['CLIN'] == clin]

		if location is not None:
			df = df.loc[df['Location'] == location]

		grouped = df.groupby(['Date', 'CLIN', 'Location', 'City', 'SubCLIN', 'Category', 'EmployeeName', 'TaskName', 'Rate', 'HourlyRateReg', 'PostingRate', 'HazardRate'], as_index=False).agg({'Hours': 'sum'})

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

		pivot['HoursReg'] = pivot['Regular'] + pivot['LocalHoliday'] + pivot['Admin']
		pivot['HoursOT'] = pivot['Overtime'] + pivot['On-callOT'] + pivot['ScheduledOT'] + pivot['UnscheduledOT']
		pivot['HoursTotal'] = pivot['HoursReg'] + pivot['HoursOT']
		pivot['RegularWages'] = pivot['Regular'] * pivot['HourlyRateReg'] # only use "Regular" hours for posting, not OT nor other type of regular hours
		pivot['Posting'] = pivot['RegularWages'] * pivot['PostingRate']
		pivot['Hazard'] = pivot['RegularWages'] * pivot['HazardRate']

		pivot = pivot[[
			'Date', 'CLIN', 'Location', 'City', 'SubCLIN', 'Category', 'EmployeeName', 
			'Regular', 'LocalHoliday', 'Holiday', 'Vacation', 'Admin', 'Bereavement', 
			'Overtime', 'On-callOT', 'ScheduledOT', 'UnscheduledOT', 
			'HoursReg', 'HoursOT', 'HoursTotal',
			'HourlyRateReg', 'RegularWages', 
			'PostingRate', 'Posting', 
			'HazardRate', 'Hazard'
		]]

		# pivot.sort_values(['Date', 'Location', 'City', 'SubCLIN', 'Category', 'EmployeeName'], inplace=True)

		return pivot
	
	def isIn(self, employeeNumberList):
		records = self.data.loc[(self.data['Number'].isin(employeeNumberList))]
		return records
	
	def startYear(self):
		return self.dateStart.strftime('%Y')
	
	def startMonth(self):
		# the month should be zero padded
		return self.dateStart.strftime('%m')
	
	def billingPeriod(self):
		startMonthName = self.dateStart.strftime('%b')
		endMonthName = self.dateEnd.strftime('%b')
		billingPeriod = f'{self.dateStart.day} {startMonthName} {self.dateStart.year} - {self.dateEnd.day} {endMonthName} {self.dateEnd.year}'
		return billingPeriod
	
	def locationsByCLIN(self):
		# DEBUG: filter out CLINs 001 and 002 to see what is left
		# print(self.data.loc[~self.data['CLIN'].isin(['001', '002'])])

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

		# debug - print the info for SubCLIN 0327
		# print(invoiceDetail.loc[invoiceDetail['SubCLIN'] == '0327'])

		invoiceDetail = invoiceDetail.groupby(['SubCLIN', 'Description', 'EmployeeName', 'Rate'], as_index=False).agg({'Hours': 'sum'})
		invoiceDetail['Amount'] = invoiceDetail['Hours'] * invoiceDetail['Rate']
		invoiceDetail = invoiceDetail[['SubCLIN', 'Description', 'EmployeeName', 'Hours', 'Rate', 'Amount']]
		invoiceDetail.sort_values(['SubCLIN', 'EmployeeName', 'Description'], ascending=[True, True, False], inplace=True)
		return invoiceDetail
	
	def groupedForCosts(self, clin=None):
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
	
	def postByCountry(self, clin=None):
		costDetail = self.details()

		if clin is not None:
			costDetail = costDetail.loc[costDetail['CLIN'] == clin]

		posts = costDetail.groupby(['Location'], as_index=False).agg({'Posting': 'sum'})
		posts['CLIN'] = '207'
		posts['Type'] = 'Post'
		posts['G&A'] = posts['Posting'] * upchargeRate
		posts['Total'] = posts['Posting'] + posts['G&A']
		posts.rename(columns={'Posting': 'Amount'}, inplace=True)
		posts = posts[['CLIN', 'Location', 'Type', 'Amount', 'G&A', 'Total']]

		hazards = costDetail.groupby(['Location', 'City'], as_index=False).agg({'Hazard': 'sum'})

		hazards['CLIN'] = '208'
		hazards['Type'] = 'Hazard'
		hazards['G&A'] = hazards['Hazard'] * upchargeRate
		hazards['Total'] = hazards['Hazard'] + hazards['G&A']
		hazards.rename(columns={'Hazard': 'Amount'}, inplace=True)
		hazards = hazards[['CLIN', 'Location', 'Type', 'Amount', 'G&A', 'Total']]

		costs = pd.concat([posts, hazards])
		costs = costs.loc[costs['Total'] > 0]
		costs.sort_values(['Location'], inplace=True)

		return costs
	
	def postSummaryByCity(self, clin=None):
		costDetail = self.details()

		if clin is not None:
			costDetail = costDetail.loc[costDetail['CLIN'] == clin]

		summary = costDetail.groupby(['Location', 'City'], as_index=False).agg({'Posting': 'sum', 'Hazard': 'sum'})
		summary['Spacer'] = ''
		summary = summary[['City', 'Location', 'Posting', 'Spacer', 'Hazard']]
		return summary

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
	
	def groupedForDetailsReport(self, clin=None, location=None):
		details = self.details(clin=clin, location=location)
		details.drop(columns=['CLIN'], inplace=True)
		# details.sort_values(['Date', 'EmployeeName'], inplace=True)

		grouped = details.groupby(['EmployeeName', 'Date'], as_index=False).agg({
			'Location': 'first',
			'City': 'first',
			'SubCLIN': 'first',
			'Category': 'first',
			'Regular': 'sum',
			'LocalHoliday': 'sum',
			'Admin': 'sum',
			'Overtime': 'sum',
			'On-callOT': 'sum',
			'ScheduledOT': 'sum',
			'UnscheduledOT': 'sum',
			'HoursReg': 'sum',
			'HoursOT': 'sum',
			'HoursTotal': 'sum',
			'HourlyRateReg': 'first',
			'RegularWages': 'sum',
			'PostingRate': 'first',
			'Posting': 'sum',
			'HazardRate': 'first',
			'Hazard': 'sum'
		})

		return grouped
	
	def groupedForPostReport(self, clin=None, location=None):
		details = self.details(clin=clin, location=location)
		details.drop(columns=['CLIN'], inplace=True)
		# details.sort_values(['Date', 'EmployeeName'], inplace=True)

		grouped = details.groupby(['SubCLIN', 'EmployeeName'], as_index=False).agg({
			'Location': 'first',
			'City': 'first',
			'Regular': 'sum',
			'HourlyRateReg': 'first',
			'RegularWages': 'sum',
			'PostingRate': 'first',
			'Posting': 'sum',
			'HazardRate': 'first',
			'Hazard': 'sum'
		})

		# reorder the columns to be more useful
		grouped = grouped[[
			'City', 'SubCLIN', 'EmployeeName',
			'Regular', 'HourlyRateReg', 'RegularWages', 
			'PostingRate', 'Posting',
			'HazardRate', 'Hazard'
		]]

		return grouped
	
	def byEmployee(self):
		grouped = self.data.groupby(['Region', 'EmployeeName', 'TaskName', 'State'], as_index=False).agg({'Hours': 'sum'})

		pivot = grouped.pivot_table(index=['Region', 'EmployeeName', 'State'], columns='TaskName', values='Hours').reset_index()

		for taskName in TaskNames.values():
			if taskName not in pivot.columns:
				pivot[taskName] = 0
			else:
				pivot[taskName] = pivot[taskName].fillna(0)
		
		pivot['HoursReg'] = pivot['Regular'] + pivot['LocalHoliday'] + pivot['Admin']
		pivot['HoursOT'] = pivot['Overtime'] + pivot['On-callOT'] + pivot['ScheduledOT'] + pivot['UnscheduledOT']
		pivot['HoursTotal'] = pivot['HoursReg'] + pivot['HoursOT']

		pivot = pivot[[
			'Region', 'EmployeeName', 'State',
			'Regular', 'LocalHoliday', 'Holiday', 'Vacation', 'Admin', 'Bereavement', 
			'Overtime', 'On-callOT', 'ScheduledOT', 'UnscheduledOT', 
			'HoursReg', 'HoursOT', 'HoursTotal'
		]]

		return pivot

	def byDate(self, clin=None, location=None):
		df = self.data.copy()

		if clin is not None:
			df = df.loc[df['CLIN'] == clin]

		if location is not None:
			df = df.loc[df['Location'] == location]

		grouped = df.groupby(['Region', 'EmployeeName', 'Date', 'TaskName', 'State'], as_index=False).agg({'Hours': 'sum'})

		pivot = grouped.pivot_table(index=['Region', 'EmployeeName', 'Date', 'State'], columns='TaskName', values='Hours').reset_index()

		for taskName in TaskNames.values():
			if taskName not in pivot.columns:
				pivot[taskName] = 0
			else:
				pivot[taskName] = pivot[taskName].fillna(0)
		
		pivot['HoursReg'] = pivot['Regular'] + pivot['LocalHoliday'] + pivot['Admin']
		pivot['HoursOT'] = pivot['Overtime'] + pivot['On-callOT'] + pivot['ScheduledOT'] + pivot['UnscheduledOT']
		pivot['HoursTotal'] = pivot['HoursReg'] + pivot['HoursOT']

		pivot = pivot[[
			'Region', 'Date', 'EmployeeName', 'State',
			'Regular', 'LocalHoliday', 'Holiday', 'Vacation', 'Admin', 'Bereavement', 
			'Overtime', 'On-callOT', 'ScheduledOT', 'UnscheduledOT', 
			'HoursReg', 'HoursOT', 'HoursTotal'
		]]

		return(pivot)

if __name__ == '__main__':
	import sys

	billingRates = BillingRates(verbose=False)

	activityFilename = sys.argv[1] if len(sys.argv) > 1 else None

	print(f'Parsing billing data from {activityFilename}')

	if activityFilename is None:
		print(f'Usage: {sys.argv[0]} <activity file>')
		exit()

	time = EmployeeTime(activityFilename, verbose=True)
	time.joinWith(billingRates)

	time.groupedForPostReport(clin='001')

	exit()

	data = time.groupedForInvoicing(clin='002', location='NATO')
	now = pd.Timestamp.now().strftime("%m%d%H%M")

	outputFile = f'Status-{time.startYear()}{time.startMonth()}-{now}.xlsx'

	timeByDate = time.byDate()
	timeByEmployee = time.byEmployee()

	with pd.ExcelWriter(outputFile) as writer:
		timeByDate.to_excel(writer, sheet_name='Date', startrow=0, startcol=0, header=True, index=False)
		timeByEmployee.to_excel(writer, sheet_name='Employee', startrow=0, startcol=0, header=True, index=False)

	# Apply formatting in place
	workbook = load_workbook(outputFile)

	for styleName in styles.keys():
		workbook.add_named_style(styles[styleName])
		
	worksheet = workbook['Date']
	formatTimeByDate(worksheet)

	worksheet = workbook['Employee']
	formatTimeByEmployee(worksheet)

	workbook.save(outputFile)