#!/usr/local/bin/python
from datetime import datetime
import pandas as pd

from BillingActivity import BillingActivity
from BillingActivityIntacct import BillingActivityIntacct

from openpyxl import load_workbook
from InvoiceStyles import styles
from InvoiceFormat import formatActivityDataTab, formatPivotTab, formatDiffsTab, highlightDiffs, formatJoinTab, formatJoinedPivotTab

def compareValues(value1, value2):
	# if any of the values are NaN, then they are not equal
	if value1.isnull().values.any() or value2.isnull().values.any():
		return False
	
	# if any of the values are empty, then they are not equal
	if value1.empty or value2.empty:
		return False
	
	# if any of the values are not floats, then they are not equal
	if not value1.dtype == float or not value2.dtype == float:
		return False
	
	# if any of the values are not equal, then they are not equal
	if not value1.equals(value2):
		return False
	
	return True

if __name__ == '__main__':
	import sys

	if len(sys.argv) < 3:
		print(f'Usage: {sys.argv[0]} <IntacctActivityFile> <TcpActivityFile>')
		exit()

	#############################################################
	# Raw details
	#############################################################
	
	intacctActivityFilename = sys.argv[1]
	intacctActivity = BillingActivityIntacct(intacctActivityFilename, verbose=False)
	intacctActivity.data.sort_values(by=['Date', 'EmployeeName', 'TaskName'], inplace=True)

	tcpActivityFilename = sys.argv[2]
	tcpActivity = BillingActivity(tcpActivityFilename, verbose=False)
	tcpActivity.data.sort_values(by=['Date', 'EmployeeName', 'TaskName'], inplace=True)

	#############################################################
	# Grouped By Date
	#############################################################
	intacctDate = intacctActivity.data.groupby(['Date']).agg({'Hours': 'sum'}).reset_index()
	tcpDate = tcpActivity.data.groupby(['Date']).agg({'Hours': 'sum'}).reset_index()

	joinedDate = intacctDate.join(tcpDate.set_index(['Date']), on=['Date'], how='left', rsuffix='_TCP')
	dateDiffs = joinedDate.loc[joinedDate['Hours'] != joinedDate['Hours_TCP']]

	#############################################################
	# TaskNames
	#############################################################

	taskNames = ['Regular', 'Admin', 'Vacation', 'Holiday', 'LocalHoliday', 'Bereavement', 'Overtime', 'On-callOT', 'ScheduledOT', 'UnscheduledOT']

	tasksSet = set(taskNames)
	intacctTasks = set(intacctActivity.data['TaskName'].unique())
	tcpTasks = set(tcpActivity.data['TaskName'].unique())

	intacctTaskDiffs =  intacctTasks - tasksSet
	tcpTaskDiffs =  tcpTasks - tasksSet

	# print(f'\nTaskNames:')
	# print(f'All: {len(taskNames)}: {taskNames}')
	
	if len(intacctTaskDiffs) > 0:
		print(f'Intacct Diffs: {len(intacctTaskDiffs)}: {intacctTaskDiffs}')

	if len(tcpTaskDiffs) > 0:
		print(f'TCP Diffs: {len(tcpTaskDiffs)}: {tcpTaskDiffs}')

	#############################################################
	# Grouped By Task
	#############################################################
	intacctTask = intacctActivity.data.groupby(['TaskName']).agg({'Hours': 'sum'}).reset_index()
	tcpTask = tcpActivity.data.groupby(['TaskName']).agg({'Hours': 'sum'}).reset_index()

	joinedTask = intacctTask.join(tcpTask.set_index(['TaskName']), on=['TaskName'], how='left', rsuffix='_TCP')
	taskDiffs = joinedTask.loc[joinedTask['Hours'] != joinedTask['Hours_TCP']]
	# print(f'\nTask Diffs ({len(taskDiffs)}):')
	# print(taskDiffs)

	#############################################################
	# Grouped By Employee
	#############################################################
	intacctEmployee = intacctActivity.data.groupby(['EmployeeName']).agg({'Hours': 'sum'}).reset_index()
	tcpEmployee = tcpActivity.data.groupby(['EmployeeName']).agg({'Hours': 'sum'}).reset_index()

	joinedEmployee = intacctEmployee.join(tcpEmployee.set_index(['EmployeeName']), on=['EmployeeName'], how='left', rsuffix='_TCP')
	employeeDiffs = joinedEmployee.loc[joinedEmployee['Hours'] != joinedEmployee['Hours_TCP']]
	# print(f'\nEmployee Diffs ({len(employeeDiffs)}):')
	# print(employeeDiffs)

	#############################################################
	# Pivot Task by Date
	#############################################################
	intacctTaskGrouped = intacctActivity.data.groupby(['Date', 'EmployeeName', 'TaskName']).agg({'Hours': 'sum'}).reset_index()

	intacctTaskPivot = intacctTaskGrouped.pivot_table(index=['Date', 'EmployeeName'], columns='TaskName', values='Hours').reset_index()
	intacctTaskPivot.fillna(0.0, inplace=True)

	for taskName in taskNames:
		if taskName not in intacctTaskPivot.columns:
			intacctTaskPivot[taskName] = 0.0
		else:
			intacctTaskPivot[taskName] = intacctTaskPivot[taskName].fillna(0.0)

	intacctTaskPivot = intacctTaskPivot[['Date', 'EmployeeName'] + list(taskNames)]
	intacctTaskPivot['Subtotal'] = intacctTaskPivot[list(taskNames)].sum(axis=1)

	tcpTaskGrouped = tcpActivity.data.groupby(['Date', 'EmployeeName', 'TaskName']).agg({'Hours': 'sum'}).reset_index()
	tcpTaskPivot = tcpTaskGrouped.pivot_table(index=['Date', 'EmployeeName'], columns='TaskName', values='Hours').reset_index()
	tcpTaskPivot.fillna(0.0, inplace=True)

	for taskName in taskNames:
		if taskName not in tcpTaskPivot.columns:
			tcpTaskPivot[taskName] = 0.0
		else:
			tcpTaskPivot[taskName] = tcpTaskPivot[taskName].fillna(0.0)
	
	tcpTaskPivot = tcpTaskPivot[['Date', 'EmployeeName'] + list(taskNames)]
	tcpTaskPivot['Subtotal'] = tcpTaskPivot[list(taskNames)].sum(axis=1)

	joinedTaskPivot = intacctTaskPivot.join(tcpTaskPivot.set_index(['Date', 'EmployeeName']), on=['Date', 'EmployeeName'], how='left', rsuffix='_TCP')
	joinedTaskPivot.fillna(0.0, inplace=True)

	# rows that have different totals
	diffTotals = joinedTaskPivot.loc[joinedTaskPivot['Subtotal'] != joinedTaskPivot['Subtotal_TCP']]

	# start with rows that have the same totals and eliminate rows do not have differences
	diffValues = joinedTaskPivot.loc[joinedTaskPivot['Subtotal'] == joinedTaskPivot['Subtotal_TCP']]

	rowsWithDifferences = set()
	for index in diffValues.index:
		for taskName in taskNames:
			if diffValues.loc[index, taskName] != diffValues.loc[index, taskName + '_TCP']:
				# print(f'\tAdding {index} because of {taskName}')
				rowsWithDifferences.add(index)

	diffValues = joinedTaskPivot.loc[joinedTaskPivot.index.isin(rowsWithDifferences)]

	#############################################################
	# Write output
	#############################################################

	now = datetime.now().strftime('%Y%m%d%H%M')
	outputFile = f'Activity-IntacctVsTCP-{now}.xlsx'

	with pd.ExcelWriter(outputFile) as writer:
		diffTotals.to_excel(writer, sheet_name='Diff-Total', index=False)
		diffValues.to_excel(writer, sheet_name='Diff-Value', index=False)
		# joinedTaskPivot.to_excel(writer, sheet_name='Joined', index=False)
		# intacctActivity.data.to_excel(writer, sheet_name='Intacct-All', index=False)
		# tcpActivity.data.to_excel(writer, sheet_name='TCP-All', index=False)

	# Apply formatting in place
	workbook = load_workbook(outputFile)

	for styleName in styles.keys():
			workbook.add_named_style(styles[styleName])

	print('Formatting Diff-Total tab')
	tab = workbook['Diff-Total']
	formatJoinedPivotTab(tab, len(taskNames) + 1)

	print('Formatting Diff-Value tab')
	tab = workbook['Diff-Value']
	formatJoinedPivotTab(tab, len(taskNames) + 1)

	# print('Formatting Joined tab')
	# tab = workbook['Joined']
	# formatJoinedPivotTab(tab, len(taskNames) + 1)

	# print('Formatting Intacct-All tab')
	# intacctTab = workbook['Intacct-All']
	# formatActivityDataTab(intacctTab)

	# print('Formatting TCP-All tab')
	# tcpTab = workbook['TCP-All']
	# formatActivityDataTab(tcpTab)

	workbook.save(outputFile)

	print(f'Wrote {outputFile}')