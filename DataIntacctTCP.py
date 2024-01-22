#!/usr/local/bin/python
# import re
from datetime import datetime
# from ast import literal_eval
import pandas as pd
# import numpy as np
# from xml.sax import ContentHandler, parse

from BillingRates import BillingRates
from BillingActivity import BillingActivity

from openpyxl import load_workbook

from BillingActivity import BillingActivity
from InvoiceStyles import styles
from InvoiceFormat import formatDebugTab, formatDiffsTab, highlightDiffs

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

			# drop columns that might be present
			if 'd' in df.columns:
				df.drop(columns=['d'], inplace=True)
			
			if 'Select' in df.columns:
				df.drop(columns=['Select'], inplace=True)

			# drop other unused columns
			df.drop(columns=['Line no.', 'Contract ID', 'Item ID', 'Fee percent', 'Price', 'Total', 'Descr', 'Contract Group ID', 'Location ID'], inplace=True)

			# rename columns
			df.rename(columns={
				'Date': 'Date',
				'Task ID': 'TaskID',
				'Task name': 'TaskName',
				'Qty': 'Hours',
				'Employee name': 'EmployeeName'
			}, inplace=True)

			# clean up values for TaskID
			df['TaskID'] = df['TaskID'].str.replace('3526', '3333')	# Admin
			df['TaskID'] = df['TaskID'].str.replace('3527', '3330')	# LocalHoliday
			df['TaskID'] = df['TaskID'].str.replace('3311', '3322')	# Regular
			df['TaskID'] = df['TaskID'].str.replace('3314', '3325')	# ScheduledOT
			df['TaskID'] = df['TaskID'].str.replace('3317', '3326')	# UnscheduledOT
			df['TaskID'] = df['TaskID'].str.replace('3313', '3324') # On-callOT

			# clean up values for TaskName
			df['TaskName'] = df['TaskName'].str.replace('Scheduled Overtime', 'ScheduledOT')
			df['TaskName'] = df['TaskName'].str.replace('Scheduled - Overtime', 'ScheduledOT')
			df['TaskName'] = df['TaskName'].str.replace('Unscheduled Overtime', 'UnscheduledOT')
			df['TaskName'] = df['TaskName'].str.replace('Unscheduled/ Emergency OT', 'UnscheduledOT')
			df['TaskName'] = df['TaskName'].str.replace('On Call- Overtime', 'On-callOT')
			df['TaskName'] = df['TaskName'].str.replace('Local Holiday', 'LocalHoliday')
			
			# We only need the date for these hours, not the time nor the specific in/out times
			# We want the date in a datetime, not a string
			df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.strftime("%m/%d/%Y")
			df['RateType'] = df['TaskID'].map(lambda x: RateTypes.get(x, 'Unknown'))

			# df.fillna('None', inplace=True)
			# df = df.replace('', np.NaN)

			df.sort_values(['Date', 'EmployeeName', 'TaskID'], ascending=[True, True, True], inplace=True)
			
			self.dateStart = pd.to_datetime(df['Date'].min(), errors="coerce")
			self.dateEnd = pd.to_datetime(df['Date'].max(), errors="coerce")
			self.data = df

			if verbose:
				print(f'\nRaw data from {filename}')
				print(df)
				print(f'\nDate range: {self.dateStart} to {self.dateEnd}')
	
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

if __name__ == '__main__':
	import sys

	if len(sys.argv) < 3:
		print(f'Usage: {sys.argv[0]} <IntacctActivityFile> <TcpActivityFile>')
		exit()

	intacctActivityFilename = sys.argv[1]
	intacct = BillingActivityIntacct(intacctActivityFilename, verbose=False)
	intacctData = intacct.data.copy()

	# clean up the names, types
	intacctData['EmployeeName'] = intacctData['EmployeeName'].apply(formatName)
	intacctData['Date'] = intacctData['Date'].astype(dtype='datetime64[ns]')
	intacctData['Hours'] = intacctData['Hours'].astype(float)

	intacctCleaned = intacctData.copy()
	
	intacctEmployees = intacctData.groupby(['EmployeeName', 'TaskName']).agg({'Hours': 'sum'}).reset_index()
	intacctPivot = intacctData.pivot_table(index=['EmployeeName'], columns='TaskName', values='Hours', aggfunc='sum').reset_index()

	for taskName in TaskNames.values():
		if taskName not in intacctPivot.columns:
			intacctPivot[taskName] = 0.0
		else:
			intacctPivot[taskName] = intacctPivot[taskName].fillna(0)

	intacctPivot['HoursReg'] = intacctPivot['Regular'] + intacctPivot['LocalHoliday'] + intacctPivot['Bereavement'] + intacctPivot['Vacation'] + intacctPivot['Admin']
	intacctPivot['HoursOT'] = intacctPivot['Overtime'] + intacctPivot['On-callOT'] + intacctPivot['ScheduledOT'] + intacctPivot['UnscheduledOT']
	intacctPivot['HoursTotal'] = intacctPivot['HoursReg'] + intacctPivot['HoursOT']

	intacctPivot = intacctPivot[[
		'EmployeeName', 
		'Regular', 'LocalHoliday', 'Bereavement', 'Vacation', 'Admin',
		'Overtime', 'On-callOT', 'ScheduledOT', 'UnscheduledOT',
		'HoursReg', 'HoursOT', 'HoursTotal'
	]]

	# print('\nIntacct data:')
	# print(intacctData.info())
	# print(intacctPivot)

	#############################################################
	# Both test
	#############################################################

	tcpActivityFilename = sys.argv[2]
	tcp = BillingActivity(tcpActivityFilename, verbose=False)
	tcpData = tcp.data.copy()

	# explicitly set the types
	tcpData['Date'] = tcpData['Date'].astype(dtype='datetime64[ns]')
	tcpData['Hours'] = tcpData['Hours'].astype(float)

	tcpCleaned = tcpData.copy()

	tcpEmployees = tcpData.groupby(['EmployeeName', 'TaskName']).agg({'Hours': 'sum'}).reset_index()
	tcpPivot = tcpData.pivot_table(index=['EmployeeName'], columns='TaskName', values='Hours', aggfunc='sum').reset_index()

	for taskName in TaskNames.values():
		if taskName not in tcpPivot.columns:
			tcpPivot[taskName] = 0.0
		else:
			tcpPivot[taskName] = tcpPivot[taskName].fillna(0)

	tcpPivot['HoursReg'] = tcpPivot['Regular'] + tcpPivot['LocalHoliday'] + tcpPivot['Bereavement'] + tcpPivot['Vacation'] + tcpPivot['Admin']
	tcpPivot['HoursOT'] = tcpPivot['Overtime'] + tcpPivot['On-callOT'] + tcpPivot['ScheduledOT'] + tcpPivot['UnscheduledOT']
	tcpPivot['HoursTotal'] = tcpPivot['HoursReg'] + tcpPivot['HoursOT']

	tcpPivot = tcpPivot[[
		'EmployeeName', 
		'Regular', 'LocalHoliday', 'Bereavement', 'Vacation', 'Admin', 
		'Overtime', 'On-callOT', 'ScheduledOT', 'UnscheduledOT', 
		'HoursReg', 'HoursOT', 'HoursTotal'
	]]

	# print('\nTCP data:')
	# print(tcpData.info())
	# print(tcpPivot)
	
	#############################################################
	# Join test
	#############################################################

	intacctGroupedData = intacctCleaned.groupby(['Date', 'EmployeeName']).agg({'TaskName': 'first', 'Hours': 'sum'}).reset_index()
	tcpGroupedData = tcpCleaned.groupby(['Date', 'EmployeeName']).agg({'TaskName': 'first', 'Hours': 'sum'}).reset_index()
	joined = intacctGroupedData.join(tcpGroupedData.set_index(['Date', 'EmployeeName']), on=['Date', 'EmployeeName'], how='left', rsuffix='_TCP')

	joined['Hours'] = joined['Hours'].fillna(0.0)
	joined['Hours_TCP'] = joined['Hours_TCP'].fillna(0.0)

	# if TaskName_TCP is missing, copy over from TaskName
	joined['TaskName_TCP'] = joined['TaskName_TCP'].fillna(joined['TaskName'])
	diffs = joined.loc[joined['Hours'] != joined['Hours_TCP']]
	
	print(diffs)

	now = datetime.now().strftime('%Y%m%d%H%M')

	outputFile = f'IntacctVsTCP-{now}.xlsx'
	with pd.ExcelWriter(outputFile) as writer:
		diffs.to_excel(writer, sheet_name='Diffs', index=False)
		joined.to_excel(writer, sheet_name='All', index=False)
		intacctPivot.to_excel(writer, sheet_name='Intacct', index=False)
		tcpPivot.to_excel(writer, sheet_name='TCP', index=False)
		

	# Apply formatting in place
	workbook = load_workbook(outputFile)

	for styleName in styles.keys():
			workbook.add_named_style(styles[styleName])
		
	worksheet = workbook['Diffs']
	# TODO: update
	# formatDiffsTab(worksheet, 'Intacct', 'TCP')

	worksheet = workbook['All']
	formatDiffsTab(worksheet, 'Intacct', 'TCP')

	intacctTab = workbook['Intacct']
	formatDebugTab(intacctTab)

	tcpTab = workbook['TCP']
	formatDebugTab(tcpTab)

	highlightDiffs(intacctTab, tcpTab)

	workbook.save(outputFile)

	print(f'Wrote {outputFile}')