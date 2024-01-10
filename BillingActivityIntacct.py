#!/usr/local/bin/python
import re
from datetime import datetime
from ast import literal_eval
import pandas as pd
import numpy as np
from xml.sax import ContentHandler, parse

from BillingRates import BillingRates

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

	# print(f'formatName({text}) -> {lastName}, {firstName} {middleInitial}')
	return f'{lastName}, {firstName} {middleInitial}'

def cleanupTaskID(text):
	if text in TaskNameMap:
		return TaskNameMap[text]
	
	return text

def cleanupTask(text):
	if text in TaskMap:
		return TaskMap[text]
	
	return text
	
class BillingActivityIntacct:
	def __init__(self, filename=None, verbose=False):
		self.data = None	# a dataframe containing the full billing information loaded from a file
		self.dateStart = datetime(1970,1,1)	# start date of the billing period loaded from a file
		self.dateEnd = datetime(3000,1,1)	# end date of the billing period loaded from a file

		if filename is not None:
			if verbose:
				print(f'Parsing billing data from {filename}')

			# we get data from Intacct in an CSV format
			df = pd.read_csv(filename, header=0, dtype=str)

			# drop unused columns
			if 'd' in df.columns:
				df.drop(columns=['d'], inplace=True)

			if 'Select' in df.columns:
				df.drop(columns=['Select'], inplace=True)

			df.drop(columns=['Line no.', 'Contract ID', 'Item ID', 'Fee percent', 'Price', 'Total', 'Descr', 'Contract Group ID', 'Location ID'], inplace=True)

			# rename columns
			df.rename(columns={
				'Date': 'Date',
				'Task ID': 'TaskID',
				'Task name': 'TaskName',
				'Qty': 'Hours',
				'Employee name': 'EmployeeName'
			}, inplace=True)

			# clean up values for TaskID, TaskName
			df['TaskID'] = df['TaskID'].apply(cleanupTask)
			df['TaskName'] = df['TaskName'].apply(cleanupTask)
			
			# clean up values for EmployeeName
			df['EmployeeName'] = df['EmployeeName'].apply(formatName)

			# We only need the date for these hours, not the time nor the specific in/out times
			# We want the date in a datetime, not a string
			df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.strftime("%m/%d/%Y")
			df['RateType'] = df['TaskID'].map(lambda x: RateTypes.get(x, 'Unknown'))

			df['Date'] = df['Date'].astype(dtype='datetime64[ns]')
			df['Hours'] = df['Hours'].astype(float)

			# df.fillna('None', inplace=True)
			# df = df.replace('', np.NaN)

			before = len(df)
			df = df.drop_duplicates(subset=['Date', 'EmployeeName', 'TaskName', 'Hours'], keep='first')
			after = len(df)

			if before != after:
				print(f'WARNING: dropped {before - after} duplicate records')

			df.sort_values(['Date', 'EmployeeName', 'TaskID'], ascending=[True, True, True], inplace=True)
			
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

		# Intacct data does not have an EmployeeID, so we need to join on the EmployeeName
		# joined = self.data.join(billingRates.data.set_index('EmployeeID'), on='Number', how='left', rsuffix='_rates')
		joined = self.data.join(billingRates.data.set_index('EmployeeName'), on='EmployeeName', how='left', rsuffix='_rates')

		joined['Rate'] = np.where(joined['RateType'] == 'Overtime', joined['BillRateOT'], joined['BillRateReg'])
		joined['Description'] = np.where(joined['RateType'] == 'Overtime', '(Overtime)', joined['Category'])
		joined['SubCLIN'] = joined['SubCLIN'].str.replace('X', baseYear)

		# reorder the columns to be more useful
		joined = joined[['Date', 'CLIN', 'Location', 'City', 'SubCLIN', 'Category', 'Description', 'EmployeeName', 'TaskID', 'TaskName', 'Hours', 'Rate', 'HourlyRateReg', 'PostingRate', 'HazardRate']]
		self.data = joined

	def details(self, clin=None):
		df = self.data.copy()
		
		if clin is not None:
			df = df.loc[df['CLIN'] == clin]

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
		pivot['Wages'] = pivot['HoursReg'] * pivot['HourlyRateReg']
		pivot['Posting'] = pivot['Wages'] * pivot['PostingRate']
		pivot['Hazard'] = pivot['Wages'] * pivot['HazardRate']

		pivot = pivot[[
			'Date', 'CLIN', 'Location', 'City', 'SubCLIN', 'Category', 'EmployeeName', 
			'Regular', 'LocalHoliday', 'Holiday', 'Vacation', 'Admin', 'Bereavement', 
			'Overtime', 'On-callOT', 'ScheduledOT', 'UnscheduledOT', 
			'HoursReg', 'HoursOT', 'HoursTotal',
			'HourlyRateReg', 'Posting', 'Hazard'
		]]

		pivot.sort_values(['Date', 'Location', 'City', 'SubCLIN', 'Category', 'EmployeeName'], inplace=True)

		return pivot

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

if __name__ == '__main__':
	import sys

	billingRates = BillingRates(verbose=False)

	activityFilename = sys.argv[1] if len(sys.argv) > 1 else None

	print(f'Parsing billing data from {activityFilename}')

	if activityFilename is None:
		print(f'Usage: {sys.argv[0]} <billing activity file>')
		exit()

	activity = BillingActivityIntacct(activityFilename, verbose=False)

	print(f'\Invoice Details:')
	print(f'Date range: {activity.dateStart} to {activity.dateEnd}')

	locationInfo = activity.locationsByCLIN()
	print(f'LocationInfo: {locationInfo}\n')

	now = pd.Timestamp.now().strftime("%Y%m%d%H%M")
	outputFile = f'BillingActivityIntacct-{now}.xlsx'
	
	with pd.ExcelWriter(outputFile) as writer:

		for clin in locationInfo.keys():
			print(f'\nLabor Invoices for CLIN: {clin}')

			for location in locationInfo[clin]:
				print(f'Invoice for: {location}')
				data = activity.groupedForInvoicing(clin=clin, location=location)
				data.to_excel(writer, sheet_name=f'Labor-{location}', index=False)
				print('\n'.join(data.to_string(index=False).split('\n')[1:]))

			print(f'\nCost Invoices for CLIN: {clin}')
			data = activity.groupedForCosts(clin=clin)
			data.to_excel(writer, sheet_name=f'Costs-{location}', index=False)
			print('\n'.join(data.to_string(index=False).split('\n')[1:]))

		print('\nDetails:')
		details = activity.details(clin=clin)
		print(details)
		details.to_excel(writer, sheet_name=f'Details-{clin}', index=False)


