#!/usr/local/bin/python
import pandas as pd
from openpyxl import load_workbook

from Config import Config
from LaborData import LaborData

from InvoiceStyles import styles
from InvoiceFormat import formatCostsTab, formatPostDetails

config = Config()
Regions = {}

versionString = f'v{config.data["version"]}'
CountryApprovers = config.data['approvers']

# get region dictionary from config and swap keys and values
# to make it easy to look up the region name from the CLIN
for region in config.data['regions']:
	Regions[config.data['regions'][region]] = region

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

	if len(sys.argv) < 2:
		print(f'Usage: {sys.argv[0]} <billing activity file>')
		sys.exit(1)

	filename = sys.argv[1]

	labor = LaborData.fromReportFile(filename)
	time = labor.time

	startYear = time.startYear()
	startMonth = time.startMonthName()
	locationInfo = time.locationsByCLIN()
	invoiceNumberValue = config.data['nextInvoiceNumber']

	for clin in locationInfo.keys():
		region = Regions[clin]
		invoiceData = labor.invoiceData[clin]

		prefix = config.data['filenamePrefixes']['postInvoices']
		pattern = f'{prefix}-{region}-{startYear}-{startMonth}'
		outputFile = f'{pattern}.xlsx'

		sheetInfo = {}
		firstRow = 3
		spaceToSummary = 4

		with pd.ExcelWriter(outputFile) as writer:
			sheetName = f'Post-{region}'

			costs = time.postByCountry(clin=clin)
			post = time.postSummaryByCity(clin=clin)
			numPostRows = post.shape[0]

			postData = time.groupedForPostReport(clin=clin)
			numPostDetailRows = postData.shape[0]
			invoiceData.addPostDetail(postData)

			hazard = time.hazardSummaryByCity(clin=clin)
			numHazardRows = hazard.shape[0]

			hazardData = time.groupedForHazardReport(clin=clin)
			invoiceData.addHazardDetail(hazardData)

			hazardDetails = time.groupedForHazardReport(clin=clin)
			numHazardDetailRows = hazardDetails.shape[0]

			invoiceAmount = costs['Total'].sum()

			if invoiceAmount > 0:
				summaryStartRow = 22
				rowsToSum = []

				rows = costs.shape[0]
				costs.to_excel(writer, sheet_name=sheetName, startrow=summaryStartRow, startcol=0, header=False)
				rowsToSum.append((summaryStartRow + 1, summaryStartRow + rows))
				summaryStartRow = summaryStartRow + rows + spaceToSummary

				invoiceNumber = f'SD-{invoiceNumberValue:04d}'
				invoiceNumberValue += 1		

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
					'hazardRows': numHazardRows,
					'hazardDetailRows': numHazardDetailRows
				}

				sheetInfo[sheetName] = invoiceDetail

			sheetName = f'Post-{region}-Post'
			postData.to_excel(writer, sheet_name=sheetName, startrow=firstRow, startcol=0, header=True, index=False)
			post.to_excel(writer, sheet_name=sheetName, startrow=numPostDetailRows + firstRow + spaceToSummary, startcol=5, header=False, index=False)

			if numHazardDetailRows > 0:
				sheetName = f'Post-{region}-Hazard'
				hazardDetails.to_excel(writer, sheet_name=sheetName, startrow=firstRow, startcol=0, header=True, index=False)
				hazard.to_excel(writer, sheet_name=sheetName, startrow=numHazardDetailRows + firstRow + spaceToSummary, startcol=5, header=False, index=False)
				
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

			if info['hazardDetailRows'] > 0:
				detailSheetName = f'Post-{region}-Hazard'
				worksheet = workbook[detailSheetName]
				postTitle = f'{region} Hazard {time.dateStart.strftime("%B")} {startYear}'
				formatPostDetails(worksheet, postTitle, firstRow, info['hazardDetailRows'], spaceToSummary, info['hazardRows'])

		workbook.save(outputFile)