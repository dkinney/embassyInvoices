#!/usr/local/bin/python
import re
from datetime import datetime
from ast import literal_eval
import pandas as pd
import numpy as np
from xml.sax import ContentHandler, parse

from BillingRates import BillingRates
from BillingActivityIntacct import BillingActivityIntacct

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
	
if __name__ == '__main__':
	import sys

	activityFilename = sys.argv[1] if len(sys.argv) > 1 else None

	if activityFilename is None:
		print(f'Usage: {sys.argv[0]} <billing activity file>')
		exit()

	activity = BillingActivityIntacct(activityFilename, verbose=True)
	data = activity.data[['Date', 'EmployeeName', 'TaskName', 'Hours']]

	compareToFilename = sys.argv[2] if len(sys.argv) > 2 else None

	if compareToFilename is not None:
		compareTo = BillingActivityIntacct(compareToFilename, verbose=False)
		compareData = compareTo.data[['Date', 'EmployeeName', 'TaskName', 'Hours']]

		both = pd.concat([data, compareData], ignore_index=True)
		both = both.drop_duplicates(keep=False)
		print(f'Unique rows in {activityFilename} compared to {compareToFilename}')
		print(both)

	else:
		print(f'\nAll Data from {activityFilename} ({len(activity.data)}):')
		print(activity.data)

		debug = activity.data.loc[activity.data['EmployeeName'] == 'Baugher James V']
		print(debug)

		dedup = activity.data.drop_duplicates(subset=['Date', 'EmployeeName', 'TaskName', 'Hours'], keep='first')
		print(f'\nDeduplicated Data from {activityFilename} ({len(dedup)}):')
		print(dedup)

		# debug = dedup.loc[activity.data['EmployeeName'] == 'Baugher James V']
		# print(debug)