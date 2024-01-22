#!/usr/local/bin/python
import pandas as pd
from openpyxl import load_workbook, Workbook

from BillingActivity import BillingActivity
from BillingActivityIntacct import BillingActivityIntacct

from InvoiceStyles import styles
from InvoiceFormat import formatEmployeesTab, formatDaysTab, formatTasksTab

if __name__ == '__main__':
	import sys

	if len(sys.argv) < 3:
		print(f'Usage: {sys.argv[0]} <IntacctActivityFile> <TcpActivityFile>')
		exit()

	intacctActivityFilename = sys.argv[1]
	intacct = BillingActivityIntacct(intacctActivityFilename, verbose=False)
	intacctData = intacct.data.copy()

	# print(intacctData.info())
	# print(intacctData)

	tcpActivityFilename = sys.argv[2]
	tcp = BillingActivity(tcpActivityFilename, verbose=False)
	tcpData = tcp.data.copy()

	# print(tcpData.info())
	# print(tcpData)

	intacctGrouped = intacctData.groupby(['Date', 'EmployeeName', 'TaskName']).agg({'Hours': 'sum'}).reset_index()
	tcpGrouped = tcpData.groupby(['Date', 'EmployeeName', 'TaskName']).agg({'Hours': 'sum'}).reset_index()

	data = intacctGrouped.join(tcpGrouped.set_index(['Date', 'EmployeeName', 'TaskName']), on=['Date', 'EmployeeName', 'TaskName'], how='left', rsuffix='_TCP')

	grouped = data.groupby(['EmployeeName']).agg({'Hours': 'sum', 'Hours_TCP': 'sum'}).reset_index()

	diffs = grouped.loc[grouped['Hours'] != grouped['Hours_TCP']]
	print(f'\nThere are differences for {len(diffs)} employees:')
	print(diffs)

	taskNames = data['TaskName'].unique()

	# print(f'\nTask names: {taskNames}')

	intacctPivot = data.pivot_table(index=[
		'Date', 'EmployeeName',
	], columns='TaskName', values='Hours').reset_index()

	for taskName in taskNames:
		if taskName not in intacctPivot.columns:
			intacctPivot[taskName] = 0
		else:
			intacctPivot[taskName] = intacctPivot[taskName].fillna(0)

	# print(f'\n\nIntacct Pivot:')
	# print(intacctPivot.info())

	# aggDetail = {}
	# for taskName in taskNames:
	# 	aggDetail[taskName] = 'sum'

	# intacctGrouped = intacctPivot.groupby(['EmployeeName']).agg(aggDetail).reset_index()
	# print(intacctGrouped)
	# boylan = intacctGrouped.loc[intacctGrouped['EmployeeName'].str.contains('Boylan')]
	# print(boylan)

	tcpPivot = data.pivot_table(index=[
		'Date', 'EmployeeName'
	], columns='TaskName', values='Hours_TCP').reset_index()

	for taskName in taskNames:
		if taskName not in tcpPivot.columns:
			tcpPivot[taskName] = 0
		else:
			tcpPivot[taskName] = tcpPivot[taskName].fillna(0)

	# print(f'\n\nTCP Pivot:')
	# print(tcpPivot.info())

	# tcpGrouped = tcpPivot.groupby(['EmployeeName']).agg(aggDetail).reset_index()
	# boylan2 = tcpGrouped.loc[tcpGrouped['EmployeeName'].str.contains('Boylan')]
	# print(boylan2)
	# exit()

	joined = intacctPivot.join(tcpPivot.set_index(['Date', 'EmployeeName']), on=['Date', 'EmployeeName'], how='left', rsuffix='_TCP')
	joined['Subtotal'] = joined[taskNames].sum(axis=1)
	joined['Subtotal_TCP'] = joined[[f'{taskName}_TCP' for taskName in taskNames]].sum(axis=1)

	joined = joined.fillna(0)

	outputFile = f'Verify-{intacct.dateStart.strftime("%Y%m")}.xlsx'

	print(f'Writing {outputFile}...')
	
	with pd.ExcelWriter(outputFile) as writer:
		grouped.to_excel(writer, sheet_name='Employees', index=False)
		data.to_excel(writer, sheet_name='Days', index=False)
		joined.to_excel(writer, sheet_name='Tasks', index=False)

	# Apply formatting in place
	workbook = load_workbook(outputFile)

	for styleName in styles.keys():
		workbook.add_named_style(styles[styleName])
		
	worksheet = workbook['Employees']
	formatEmployeesTab(worksheet)

	worksheet = workbook['Days']
	formatDaysTab(worksheet)

	worksheet = workbook['Tasks']
	formatTasksTab(worksheet)

	workbook.save(outputFile)