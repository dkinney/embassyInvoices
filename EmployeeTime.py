#!/usr/local/bin/python
import sys
import os
import glob
import re
from datetime import datetime
from ast import literal_eval
import pandas as pd
import numpy as np
from xml.sax import ContentHandler, parse

from EmployeeInfo import EmployeeInfo
from BillingRates import BillingRates
from Allowances import Allowances

from openpyxl import load_workbook
from InvoiceStyles import styles
from InvoiceFormat import formatTimeByDate, formatTimeByEmployee

from Config import Config
config = Config()

contractNumber = config.data['contractNumber']
baseYear = config.data['baseYear']
upchargeRate = config.data['upchargeRate']
# clins = config.data['regions']

# get region dictionary from config and swap keys and values
# to make it easy to look up the region name from the CLIN
Regions = {}
for region in config.data['regions']:
	Regions[config.data['regions'][region]] = region

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

	return f'{lastName}, {firstName}'

def cleanupTask(text):
	if text in TaskMap:
		return TaskMap[text]
	
	return text

def rate(row):
	if row['RateType'] == 'Overtime':
		return row['BillRateOT']
	
	if row['RateType'] == 'Regular':
		return row['BillRateReg']
	
	return 0

def description(row):
	if row['RateType'] == 'Overtime':
		return '(Overtime)'
	
	if row['RateType'] == 'Regular':
		return row['Category']
	
	return 0	

class EmployeeTime:
	def __init__(self, filename=None, verbose=False):
		self.data = None	# a dataframe containing the full billing information loaded from a file
		self.dateStart = datetime(1970,1,1)	# start date of the billing period loaded from a file
		self.dateEnd = datetime(3000,1,1)	# end date of the billing period loaded from a file

		if filename is not None:
			if verbose:
				print(f'Parsing activity data from {filename}')
			
			# read from csv in latin1 encoding
			df = pd.read_csv(filename, encoding='latin1')
				
			# df = pd.read_csv(filename, converters=converters)
			df.columns = ['EmployeeName', 'EmployeeID', 'Date', 'Description', 'TaskName', 'Hours', 'State']

			# fill down the missing EmployeeName values
			df['EmployeeName'] = df['EmployeeName'].fillna(method='ffill')

			# remove rows where the Hours are missing
			df = df.dropna(subset=['Hours'])

			# we only care about the rows that start with our contract number in the Description
			df = df.loc[df['Description'].str.startswith('19AQMM23C0047')]

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
			
	def joinWith(self, employeeInfo):
		if employeeInfo.data is None:
			# nothing to do
			return

		joined = self.data.join(employeeInfo.data.set_index('EmployeeID'), on='EmployeeID', how='left', rsuffix='_info')

		unjoinedTime = joined.loc[joined['Country'].isna()]
		if len(unjoinedTime) > 0:
			print(f'{len(unjoinedTime)} records do not have a location:')
			print(unjoinedTime)
		
		# zero pad the CLIN to be 3 digit string
		joined['CLIN'] = joined['CLIN'].str.zfill(3)

		# lookup the Region from the CLIN
		joined['Region'] = joined['CLIN'].map(Regions)

		joined['Country'] = joined['Country'].fillna('Unknown')
		joined['Rate'] = joined.apply(rate, axis=1)
		joined['Rate'] = pd.to_numeric(joined['Rate'], errors="coerce")
		joined['Description'] = joined.apply(description, axis=1)
		joined['RoleID'] = joined['RoleID'].str.replace('X', baseYear)
		
		# reorder the columns to be more useful
		joined = joined[['Date', 'CLIN', 'Region', 'Country', 'PostName', 'RoleID', 'Category', 'Description', 'EmployeeName', 'TaskName', 'Hours', 'State', 'Rate', 'HourlyRate', 'PostingRate', 'DangerRate']]
	
		self.data = joined

	def details(self, clin=None, location=None):
		df = self.data.copy()
		
		if clin is not None:
			df = df.loc[df['CLIN'] == clin]

		if location is not None:
			df = df.loc[df['Country'] == location]

		grouped = df.groupby(['Date', 'CLIN', 'Country', 'PostName', 'RoleID', 'Category', 'EmployeeName', 'TaskName', 'Rate', 'HourlyRate', 'PostingRate', 'DangerRate'], as_index=False).agg({'Hours': 'sum'})

		pivot = grouped.pivot_table(index=[
			'Date', 'CLIN', 'Country', 'PostName', 'RoleID', 'Category', 'EmployeeName', 
			'Rate', 'HourlyRate', 'PostingRate', 'DangerRate'
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
		pivot['RegularWages'] = pivot['Regular'] * pivot['HourlyRate'] # only use "Regular" hours for posting, not OT nor other type of regular hours
		pivot['Posting'] = pivot['RegularWages'] * pivot['PostingRate']
		pivot['Danger'] = pivot['RegularWages'] * pivot['DangerRate']

		pivot = pivot[[
			'Date', 'CLIN', 'Country', 'PostName', 'RoleID', 'Category', 'EmployeeName', 
			'Regular', 'LocalHoliday', 'Holiday', 'Vacation', 'Admin', 'Bereavement', 
			'Overtime', 'On-callOT', 'ScheduledOT', 'UnscheduledOT', 
			'HoursReg', 'HoursOT', 'HoursTotal',
			'HourlyRate', 'RegularWages', 
			'PostingRate', 'Posting', 
			'DangerRate', 'Danger'
		]]

		return pivot
	
	def isIn(self, employeeNumberList):
		records = self.data.loc[(self.data['EmployeeID'].isin(employeeNumberList))]
		return records
	
	def startYear(self):
		return self.dateStart.strftime('%Y')
	
	def startMonth(self):
		# the month should be zero padded
		return self.dateStart.strftime('%m')
	
	def startMonthName(self):
		return self.dateStart.strftime('%b')
	
	def billingPeriod(self):
		startMonthName = self.dateStart.strftime('%b')
		endMonthName = self.dateEnd.strftime('%b')
		billingPeriod = f'{self.dateStart.day} {startMonthName} {self.dateStart.year} - {self.dateEnd.day} {endMonthName} {self.dateEnd.year}'
		return billingPeriod
	
	def locationsByCLIN(self):
		noLocation = self.data.loc[self.data['Country'].isna()]

		if len(noLocation) > 0:
			print(f'\n{len(noLocation)} records do not have a location:')
			print(noLocation)

		result = {}
		for clin in self.data['CLIN'].unique():
			clinData = self.data.loc[self.data['CLIN'] == clin]

			locations = []
			for location in clinData['Country'].unique():
				locations.append(location)

			result[clin] = locations

		return result
	
	def groupedForInvoicing(self, clin=None, location=None):
		invoiceDetail = self.data.copy()

		if clin is not None:
			invoiceDetail = invoiceDetail.loc[invoiceDetail['CLIN'] == clin]

		if location is not None:
			invoiceDetail = invoiceDetail.loc[invoiceDetail['Country'] == location]

		# NOTE: ONLY uses Approved hours
		omittedDetails = invoiceDetail.loc[self.data['State'] != 'Approved']

		if len(omittedDetails) > 0:
			print(f'\nHours omitted because of state for CLIN: {clin}, {location}: {len(omittedDetails)}')
			omittedGrouped = omittedDetails.groupby(['State'], as_index=False).agg({'Hours': 'sum'})
			print(omittedGrouped.to_string(index=False, header=False))

		invoiceDetail = invoiceDetail.loc[invoiceDetail['State'] == 'Approved']
		invoiceDetail = invoiceDetail.groupby(['RoleID', 'Description', 'EmployeeName', 'Rate'], as_index=False).agg({'Hours': 'sum'})
		invoiceDetail['Amount'] = invoiceDetail['Hours'] * invoiceDetail['Rate']
		invoiceDetail = invoiceDetail[['RoleID', 'Description', 'EmployeeName', 'Hours', 'Rate', 'Amount']]
		invoiceDetail.sort_values(['RoleID', 'EmployeeName', 'Description'], ascending=[True, True, False], inplace=True)

		return invoiceDetail
	
	def postByCountry(self, clin=None):
		costDetail = self.details()

		if clin is not None:
			costDetail = costDetail.loc[costDetail['CLIN'] == clin]

		posts = costDetail.groupby(['Country'], as_index=False).agg({'Posting': 'sum'})
		posts['CLIN'] = '207'
		posts['Type'] = 'Post'
		posts['G&A'] = posts['Posting'] * upchargeRate
		posts['Total'] = posts['Posting'] + posts['G&A']
		posts.rename(columns={'Posting': 'Amount'}, inplace=True)
		posts = posts[['CLIN', 'Country', 'Type', 'Amount', 'G&A', 'Total']]

		dangers = costDetail.groupby(['Country', 'PostName'], as_index=False).agg({'Danger': 'sum'})

		dangers['CLIN'] = '208'
		dangers['Type'] = 'Danger'
		dangers['G&A'] = dangers['Danger'] * upchargeRate
		dangers['Total'] = dangers['Danger'] + dangers['G&A']
		dangers.rename(columns={'Danger': 'Amount'}, inplace=True)
		dangers = dangers[['CLIN', 'Country', 'Type', 'Amount', 'G&A', 'Total']]

		costs = pd.concat([posts, dangers])
		costs = costs.loc[costs['Total'] > 0]
		costs.sort_values(['Country'], inplace=True)

		return costs
	
	def postSummaryByPostName(self, clin=None):
		costDetail = self.details()

		if clin is not None:
			costDetail = costDetail.loc[costDetail['CLIN'] == clin]

		summary = costDetail.groupby(['Country', 'PostName'], as_index=False).agg({'Posting': 'sum'})
		return summary

	def dangerPaySummaryByPostName(self, clin=None):
		costDetail = self.details()

		if clin is not None:
			costDetail = costDetail.loc[costDetail['CLIN'] == clin]

		summary = costDetail.groupby(['Country', 'PostName'], as_index=False).agg({'Danger': 'sum'})
		return summary

	def groupedForHoursReport(self, clin=None, location=None):
		data = self.data.copy()

		detail = self.data.copy()

		if clin is not None:
			detail = detail.loc[detail['CLIN'] == clin]

		if location is not None:
			detail = detail.loc[detail['Country'] == location]

		detail = detail.groupby(['RoleID', 'Description', 'EmployeeName', 'Rate'], as_index=False).agg({'Hours': 'sum'})
		detail['Amount'] = detail['Hours'] * detail['Rate']
		detail = detail[['RoleID', 'Description', 'EmployeeName', 'Hours', 'Rate', 'Amount']]
		detail.sort_values(['RoleID', 'EmployeeName', 'Description'], ascending=[True, True, False], inplace=True)

		return detail
	
	def groupedForDetailsReport(self, clin=None, location=None):
		details = self.details(clin=clin, location=location)

		# print(f'debug groupedForDetailsReport: for {clin}, {location}: {details.PostName.unique()}')

		details.drop(columns=['CLIN'], inplace=True)
		# details.sort_values(['Date', 'EmployeeName'], inplace=True)

		grouped = details.groupby(['EmployeeName', 'Date'], as_index=False).agg({
			'Country': 'first',
			'PostName': 'first',
			'RoleID': 'first',
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
			'HourlyRate': 'first',
			'RegularWages': 'sum',
			'PostingRate': 'first',
			'Posting': 'sum',
			'DangerRate': 'first',
			'Danger': 'sum'
		})

		return grouped
	
	def groupedForPostReport(self, clin=None, location=None):
		details = self.details(clin=clin, location=location)

		# print(f'debug groupedForPostReport: for {clin}, {location}: {details.PostName.unique()}')

		details.drop(columns=['CLIN'], inplace=True)
		# details.sort_values(['Date', 'EmployeeName'], inplace=True)

		grouped = details.groupby(['RoleID', 'EmployeeName'], as_index=False).agg({
			'Country': 'first',
			'PostName': 'first',
			'Regular': 'sum',
			'HourlyRate': 'first',
			'RegularWages': 'sum',
			'PostingRate': 'first',
			'Posting': 'sum'
		})

		# reorder the columns to be more useful
		grouped = grouped[[
			'PostName', 'RoleID', 'EmployeeName',
			'Regular', 'HourlyRate', 'RegularWages', 
			'PostingRate', 'Posting'
		]]

		# rename columns to be more readable
		grouped.rename(columns={
			'RoleID': 'CLIN',
			'EmployeeName': 'Name',
			'HourlyRate': 'Rate',
			'RegularWages': 'Regular Wages',
			'PostingRate': 'Post Rate',
			'Posting': 'Post Pay'
		}, inplace=True)

		return grouped
	
	def groupedForDangerReport(self, clin=None, location=None):
		details = self.details(clin=clin, location=location)
		details.drop(columns=['CLIN'], inplace=True)

		# print(f'debug groupedForDangerReport: for {clin}, {location}: {details.PostName.unique()}')

		grouped = details.groupby(['RoleID', 'EmployeeName'], as_index=False).agg({
			'Country': 'first',
			'PostName': 'first',
			'Regular': 'sum',
			'HourlyRate': 'first',
			'RegularWages': 'sum',
			'DangerRate': 'first',
			'Danger': 'sum'
		})

		# reorder the columns to be more useful
		grouped = grouped[[
			'PostName', 'RoleID', 'EmployeeName',
			'Regular', 'HourlyRate', 'RegularWages', 
			'DangerRate', 'Danger'
		]]

		# rename columns to be more readable
		grouped.rename(columns={
			'RoleID': 'CLIN',
			'EmployeeName': 'Name',
			'HourlyRate': 'Rate',
			'RegularWages': 'Regular Wages',
			'DangerRate': 'Danger Rate',
			'Danger': 'Danger Pay'
		}, inplace=True)

		# drop rows where the Danger is zero
		grouped = grouped.loc[grouped['Danger Pay'] > 0]

		return grouped
	
	# for status report
	def byEmployee(self, clin=None, location=None):
		df = self.data.copy()

		if clin is not None:
			df = df.loc[df['CLIN'] == clin]

		if location is not None:
			df = df.loc[df['Country'] == location]

		# print('byEmployee Countries: ', df['Country'].unique())

		grouped = df.groupby(['Region', 'PostName', 'EmployeeName', 'RoleID', 'TaskName', 'State'], as_index=False).agg({'Hours': 'sum'})

		pivot = grouped.pivot_table(index=['Region', 'PostName', 'EmployeeName', 'RoleID', 'State'], columns='TaskName', values='Hours').reset_index()

		for taskName in TaskNames.values():
			if taskName not in pivot.columns:
				pivot[taskName] = 0
			else:
				pivot[taskName] = pivot[taskName].fillna(0)
		
		pivot['HoursReg'] = pivot['Regular'] + pivot['LocalHoliday'] + pivot['Admin']
		pivot['HoursOT'] = pivot['Overtime'] + pivot['On-callOT'] + pivot['ScheduledOT'] + pivot['UnscheduledOT']
		pivot['HoursTotal'] = pivot['HoursReg'] + pivot['HoursOT']

		pivot = pivot[[
			'PostName', 'RoleID', 'EmployeeName', 'State',
			'Regular', 'LocalHoliday', 'Admin', 
			'Holiday', 'Vacation', 'Bereavement', 
			'Overtime', 'On-callOT', 'ScheduledOT', 'UnscheduledOT', 
			'HoursReg', 'HoursOT', 'HoursTotal'
		]]

		return pivot

	def byDate(self, clin=None, location=None):
		df = self.data.copy()

		if clin is not None:
			df = df.loc[df['CLIN'] == clin]

		if location is not None:
			df = df.loc[df['Country'] == location]

		# print('byDate Countries: ', df['Country'].unique())

		grouped = df.groupby(['Region', 'EmployeeName', 'RoleID', 'Date', 'TaskName', 'State'], as_index=False).agg({'Hours': 'sum'})

		pivot = grouped.pivot_table(index=['Region', 'EmployeeName', 'RoleID', 'Date', 'State'], columns='TaskName', values='Hours').reset_index()

		for taskName in TaskNames.values():
			if taskName not in pivot.columns:
				pivot[taskName] = 0
			else:
				pivot[taskName] = pivot[taskName].fillna(0)
		
		pivot['HoursReg'] = pivot['Regular'] + pivot['LocalHoliday'] + pivot['Admin']
		pivot['HoursOT'] = pivot['Overtime'] + pivot['On-callOT'] + pivot['ScheduledOT'] + pivot['UnscheduledOT']
		pivot['HoursTotal'] = pivot['HoursReg'] + pivot['HoursOT']

		pivot = pivot[[
			'Date', 'EmployeeName',
			'Regular', 'LocalHoliday', 'Admin', 
			'Overtime', 'On-callOT', 'ScheduledOT', 'UnscheduledOT', 
			'HoursTotal'
		]]

		pivot.rename(columns={
			'EmployeeName': 'Name',
			'On-callOT': 'On-call OT',
			'ScheduledOT': 'Sched OT',
			'UnscheduledOT': 'Unschd OT',
			'LocalHoliday': 'Local Hol',
			'HoursTotal': 'Subtotal'
		}, inplace=True)

		return(pivot)
	
	# for status report
	def statusByDate(self, clin=None, location=None):
		df = self.data.copy()

		if clin is not None:
			df = df.loc[df['CLIN'] == clin]

		if location is not None:
			df = df.loc[df['Country'] == location]

		# print('statusByDate Countries: ', df['Country'].unique())

		grouped = df.groupby(['Region', 'PostName', 'EmployeeName', 'RoleID', 'Date', 'TaskName', 'State'], as_index=False).agg({'Hours': 'sum'})

		pivot = grouped.pivot_table(index=['Region', 'PostName', 'EmployeeName', 'RoleID', 'Date', 'State'], columns='TaskName', values='Hours').reset_index()

		for taskName in TaskNames.values():
			if taskName not in pivot.columns:
				pivot[taskName] = 0
			else:
				pivot[taskName] = pivot[taskName].fillna(0)
		
		pivot['HoursReg'] = pivot['Regular'] + pivot['LocalHoliday'] + pivot['Admin']
		pivot['HoursOT'] = pivot['Overtime'] + pivot['On-callOT'] + pivot['ScheduledOT'] + pivot['UnscheduledOT']
		pivot['HoursTotal'] = pivot['HoursReg'] + pivot['HoursOT']

		pivot = pivot[[
			'PostName', 'Date', 'EmployeeName', 'RoleID', 'State',
			'Regular', 'LocalHoliday', 'Admin', 
			'Holiday', 'Vacation', 'Bereavement', 
			'Overtime', 'On-callOT', 'ScheduledOT', 'UnscheduledOT', 
			'HoursReg', 'HoursOT', 'HoursTotal'
		]]

		return(pivot)
	
	def dateDetails(self, clin:str = None, location:str = None) -> pd.DataFrame:
		regionDate = self.statusByDate(clin=clin, location=location)
		regionDate.drop(columns=['PostName'], inplace=True)
		notApproved = regionDate['State'].ne('Approved')

		if notApproved.any():
			print(f'Warning: {notApproved.sum()} hours not approved for {region} CLIN {clin}')
			print(regionDate.loc[notApproved])

		regionDate.sort_values(['EmployeeName', 'Date'], ascending=[True, True], inplace=True)

		# drop columns that are not necessary for Approvals
		regionDate.drop(columns=[
			'RoleID',
			'State',
			'Holiday',
			'Vacation',
			'Bereavement',
			'HoursReg',
			'HoursOT'
		], inplace=True)

		# rename columns for clarity
		regionDate.rename(columns={
			'EmployeeName': 'Name',
			'On-callOT': 'On-call OT',
			'ScheduledOT': 'Sched OT',
			'UnscheduledOT': 'Unschd OT',
			'LocalHoliday': 'Local Hol',
			'HoursTotal': 'Subtotal'
		}, inplace=True)

		return regionDate

	def employeeDetails(self, clin:str = None, location:str = None) -> pd.DataFrame:
		regionEmployee = self.byEmployee(clin=clin, location=location)

		# drop columns that are not necessary for Approvals
		regionEmployee.drop(columns=[
			'State',
			'Holiday',
			'Vacation',
			'Bereavement',
			'HoursReg',
			'HoursOT'
		], inplace=True)

		# rename columns for clarity
		regionEmployee.rename(columns={
			'EmployeeName': 'Name',
			'On-callOT': 'On-call OT',
			'ScheduledOT': 'Sched OT',
			'UnscheduledOT': 'Unschd OT',
			'LocalHoliday': 'Local Hol',
			'HoursTotal': 'Subtotal'
		}, inplace=True)

		return regionEmployee
	
def getUniquifier(pattern, type=None, region=None, year=None, monthName=None):
	# look for previous instances of the status file
	patternValues = re.findall(r'(.*)-(.*)-(.*)-(.*)', pattern)
	patternType = patternValues[0][0]
	patterRegion = patternValues[0][1]
	patternYear = patternValues[0][2]
	patternMonthhName = patternValues[0][3]

	statusFiles = glob.glob(pattern + '*.xlsx')

	version = 0
	for file in statusFiles:
		for vals in re.findall(r'(.*)-(.*)-(.*)-(.*)-(\d+)', file):
			typeMatches = type is None or type is not None and vals[0] == type
			regionMatches = region is None or region is not None and vals[1] == region
			yearMatches = year is None or year is not None and vals[2] == year
			monthMatches = monthName is None or monthName is not None and vals[3] == monthName

			if typeMatches and regionMatches and yearMatches and monthMatches:
				thisVersion = int(vals[len(vals)-1])
				version = max(version, thisVersion)

	version += 1
	return version

if __name__ == '__main__':
	activityFilename = sys.argv[1] if len(sys.argv) > 1 else None

	if activityFilename is None:
		print(f'Usage: {sys.argv[0]} <activity file>')
		exit()

	time = EmployeeTime(filename=activityFilename)
	effectiveDate = time.dateEnd
	allowances = Allowances(effectiveDate=effectiveDate)
	billingRates = BillingRates(effectiveDate=effectiveDate)
	billingRates.joinWith(allowances)
	employees = EmployeeInfo()
	employees.joinWith(billingRates)
	time.joinWith(employees)

	print(f'\nActivity from {time.dateStart} to {time.dateEnd}')
	# now = pd.Timestamp.now().strftime("%m%d%H%M")


	timeByDate = time.statusByDate()
	timeByDate.sort_values(['Date', 'EmployeeName'], ascending=[False, True], inplace=True)

	timeByEmployee = time.byEmployee()

	for clin in sorted(time.data.CLIN.unique()):
		region = Regions[clin]
		regionDate = time.statusByDate(clin=clin)
		regionDate.sort_values(['Date', 'EmployeeName'], ascending=[False, True], inplace=True)

		notApproved = regionDate.loc[regionDate['State'] != 'Approved']
		status = notApproved.groupby(['EmployeeName', 'State'], as_index=False).agg({'HoursTotal': 'sum'})
		hoursNotApproved = status.loc[status['State'] != 'Approved']['HoursTotal'].sum()

		if hoursNotApproved > 0:
			print(f'{Regions[clin]} ({clin}): {hoursNotApproved} Hours not approved across {len(status)} employees')

		regionEmployee = time.byEmployee(clin=clin)

		print(regionEmployee)