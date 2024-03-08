#!/usr/local/bin/python
import pandas as pd
from openpyxl import load_workbook

from Config import Config
from LaborData import LaborData

from InvoiceStyles import styles
from InvoiceFormat import formatCostsTab, formatPostDetails

config = Config()
Regions = config.data['regions']

versionString = f'v{config.data["version"]}'
CountryApprovers = config.data['approvers']

# construct a summary dataframe for writing into the file
def summaryDataframe(description:str, hours:float, amount:float) -> pd.DataFrame:
	data = {
		'blank1': '',
		'Description': description,
		'blank2': '',
		'Hours': hours,
		'blank3': '',
		'Amount': amount
	}

	return pd.DataFrame(data, index=[0])

if __name__ == '__main__':
	import sys

	if len(sys.argv) < 4:
		print(f'Usage: {sys.argv[0]} <regionName> <invoiceNumber> <billing activity file>')
		sys.exit(1)

	region = sys.argv[1]
	invoiceNumberInput = sys.argv[2]	# this is the number part, although it can have an R if a revision of a revision
	filename = sys.argv[3]

	# the invoiceNumberValue is the number part of the invoice number (before any R suffix)
	try:
		invoiceNumberValue = int(invoiceNumberInput.split('R')[0])
	except ValueError:
		print(f'Error: {invoiceNumberInput} is not a valid invoice number')
		print(f'It must be a number with zero or more R suffixes (for revisions)')
		sys.exit(1)

	# the number of R's is the revision count
	revisionCount = invoiceNumberInput.count('R') + 1

	invoiceNumberSuffix = 'R' * revisionCount
	invoiceNumber = f'SD-{invoiceNumberValue:04d}{invoiceNumberSuffix}'

	print(f'This is revision: {revisionCount}')
	print(f'Invoice Number: {invoiceNumber}')

	labor = LaborData.fromReportFile(filename)
	time = labor.time

	startYear = time.startYear()
	startMonth = time.startMonthName()
	locationInfo = time.locationsByCLIN()

	if region not in Regions.keys():
		print(f'Error: "{region}" is not a valid region name')
		print(f'Available regions are: {", ".join(Regions.keys())}')
		exit(-1)

	clin = Regions[region]
	invoiceData = labor.invoiceData[clin]

	prefix = config.data['filenamePrefixes']['revisedPost']
	pattern = f'{prefix}-{region}-{startYear}-{startMonth}'
	outputFile = f'{pattern}.xlsx'

	sheetInfo = {}
	firstRow = 3
	spaceToSummary = 4

	with pd.ExcelWriter(outputFile) as writer:
		sheetName = f'Post-{region}'

		costs = time.postByCountry(clin=clin)
		post = time.postSummaryByPostName(clin=clin)
		numPostRows = post.shape[0]

		postData = time.groupedForPostReport(clin=clin)
		numPostDetailRows = postData.shape[0]
		invoiceData.addPostDetail(postData)

		dangerPay = time.dangerPaySummaryByPostName(clin=clin)
		numDangerPayRows = dangerPay.shape[0]

		dangerPayData = time.groupedForDangerReport(clin=clin)
		invoiceData.addDangerPayDetail(dangerPayData)

		dangerPayDetails = time.groupedForDangerReport(clin=clin)
		numDangerPayDetailRows = dangerPayDetails.shape[0]

		invoiceAmount = costs['Total'].sum()

		if invoiceAmount > 0:
			summaryStartRow = 22
			rowsToSum = []

			rows = costs.shape[0]
			costs.to_excel(writer, sheet_name=sheetName, startrow=summaryStartRow, startcol=0, header=False)
			rowsToSum.append((summaryStartRow + 1, summaryStartRow + rows))
			summaryStartRow = summaryStartRow + rows + spaceToSummary

			startMonthName = time.dateStart.strftime('%b')
			endMonthName = time.dateEnd.strftime('%b')
			billingPeriod = time.billingPeriod()
			
			invoiceDetail = {
				'filename': outputFile,
				'type': 'Post',
				'invoiceNumber': invoiceNumber,
				'taskOrder': f'Post-{region}',
				'billingPeriod': billingPeriod,
				'invoiceAmount': invoiceAmount,
				'rowsToSum': rowsToSum, 
				'postRows': numPostRows,
				'postDetailRows': numPostDetailRows,
				'dangerPayRows': numDangerPayRows,
				'dangerPayDetailRows': numDangerPayDetailRows
			}

			sheetInfo[sheetName] = invoiceDetail

		sheetName = f'Post-{region}-Post'
		postData.to_excel(writer, sheet_name=sheetName, startrow=firstRow, startcol=0, header=True, index=False)
		post.to_excel(writer, sheet_name=sheetName, startrow=numPostDetailRows + firstRow + spaceToSummary, startcol=5, header=False, index=False)

		if numDangerPayDetailRows > 0:
			sheetName = f'Post-{region}-DangerPay'
			dangerPayDetails.to_excel(writer, sheet_name=sheetName, startrow=firstRow, startcol=0, header=True, index=False)
			dangerPay.to_excel(writer, sheet_name=sheetName, startrow=numDangerPayDetailRows + firstRow + spaceToSummary, startcol=5, header=False, index=False)
			
	workbook = load_workbook(outputFile)

	for styleName in styles.keys():
		workbook.add_named_style(styles[styleName])
	
	for key in sheetInfo.keys():
		worksheet = workbook[key]
		info = sheetInfo[key]
		formatCostsTab(worksheet, info)

		detailSheetName = f'Post-{region}-Post'
		worksheet = workbook[detailSheetName]
		postTitle = f'{region} Post {time.dateStart.strftime("%B")} {startYear}'
		formatPostDetails(worksheet, postTitle, firstRow, info['postDetailRows'], spaceToSummary, info['postRows'])

		if info['dangerPayDetailRows'] > 0:
			detailSheetName = f'Post-{region}-DangerPay'
			worksheet = workbook[detailSheetName]
			postTitle = f'{region} Danger Pay {time.dateStart.strftime("%B")} {startYear}'
			formatPostDetails(worksheet, postTitle, firstRow, info['dangerPayDetailRows'], spaceToSummary, info['dangerPayRows'])

	workbook.save(outputFile)