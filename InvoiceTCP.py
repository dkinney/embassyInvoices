#!/usr/local/bin/python
import pandas as pd
import numpy as np

from BillingActivity import BillingActivity

if __name__ == '__main__':
	import sys

	activityFilename = sys.argv[1] if len(sys.argv) > 1 else None

	if activityFilename is None:
		print(f'Usage: {sys.argv[0]} <billing activity file>')
		exit()

	activity = BillingActivity(activityFilename, verbose=False)
	data = activity.data.copy()

	print('------------------')
	print('Activity:')
	print(data.info())
	print(data)
	print('------------------')

	activityYear = activity.startYear()
	activityMonth = activity.startMonth()

	# we are going to create 6 files; 3 for each region for 
	# "Labor" for Labor invoices, one for each "Location" (country)
	# "Cost" for ODC invoices, one for each "Location" (country)
	# "Details" for a sorting and filtering to check the accuracy of the invoices

	# There will be a package created for each region with the 5 following files:
	# 1. Labor-YYYYMM-[RegionName]-[processingVersion].pdf
	# 2. Labor-YYYYMM-[RegionName]-[processingVersion].xlsx
	# 3. Costs-YYYYMM-[RegionName]-[processingVersion].pdf
	# 4. Costs-YYYYMM-[RegionName]-[processingVersion].xlsx
	# 5. Details-YYYYMM-[RegionName]-[processingVersion].xlsx

	# The invoice number will take the form of [InvoiceType]-[YYYYMM][CountryCode].
	# The "InvoiceType" will be "SDEL" for labor invoices and "SDEC" for ODC invoices
	laborInvoiceNumber = 'SDEL-{activityYear}{activityMonth}'
	costInvoiceNumber = 'SDEC-{activityYear}{activityMonth}'

	# The number of invoices for a region is dependent on the number of locations in that region.

	for clin in data['CLIN'].unique():
		region = Regions[clin]

		# get a list of locations for this CLIN
		locations = data[data['CLIN'] == clin]['Location'].unique()

		clinData = data[data['CLIN'] == clin]
		clinData = clinData[['Location', 'SubCLIN', 'Name', 'Hours', 'Rate', 'Amount']]

		############################################################################################################
		# create the Labor file
		############################################################################################################

		outputFile = f'Labor-{activityYear}{activityMonth}-{region}-v3.xlsx'

		with pd.ExcelWriter(outputFile) as writer:
			if verbose:
				print(f'\nCreating {outputFile}...')

			# prepare a dictionary of information for the formatting function
			sheetInfo = {}

			for location in locations:
				sheetName = f'Labor-{location}'
				locationData = data[data['Location'] == location]

				# create a summary for the bottom of the invoice
				locationSummary = locationData.groupby(['Location'], as_index=False).agg({'SubCLIN': 'first', 'Category': 'first', 'Name': 'first', 'Hours': 'sum', 'Rate': 'first', 'Amount': 'sum'})
				locationSummary = locationSummary[['SubCLIN', 'Category', 'Name', 'Hours', 'Rate', 'Amount']]

				# empty some of the columns for the summary
				locationSummary['SubCLIN'] = ' '
				locationSummary['Category'] = f'Totals for {location}'
				locationSummary['Name'] = ' '
				locationSummary['Rate'] = ' '
				
				info = locationData[['SubCLIN', 'Category', 'Name', 'Hours', 'Rate', 'Amount']]

				summaryStartRow = 22
				rowsToSum = []

				for subCLIN in info['SubCLIN'].unique():
					clinData = info[info['SubCLIN'] == subCLIN]
					rows = clinData.shape[0]
					clinData.to_excel(writer, sheet_name=sheetName, startrow=summaryStartRow, startcol=0, header=False)
					rowsToSum.append((summaryStartRow + 1, summaryStartRow + rows))
					summaryStartRow = summaryStartRow + rows + 2

				sheetInfo[sheetName] = {
					'invoiceNumber': laborInvoiceNumber + CountryCodes[location],
					'taskOrder': location,
					'dateStart': billing.dateStart.strftime("%m/%d/%Y"),
					'dateEnd': billing.dateEnd.strftime("%m/%d/%Y"),
					'invoiceAmount': locationData['Amount'].sum(),
					'rowsToSum': rowsToSum
				}

				locationSummary.to_excel(writer, sheet_name=sheetName, startrow=summaryStartRow, startcol=0, header=False)

		# Apply formatting in place
		workbook = load_workbook(outputFile)

		for styleName in styles.keys():
			workbook.add_named_style(styles[styleName])

		for location in locations:
			key = f'Labor-{location}'
			worksheet = workbook[key]
			info = sheetInfo[key]
			formatInvoiceTab(worksheet, info)

		workbook.save(outputFile)
		print(f'Invoices available at {outputFile}')

		############################################################################################################
		# Gather the details, used for both Costs and Details
		############################################################################################################
		
		byEmployee = billing.byEmployee(fullDetail=True)
		detailData = byEmployee.join(employeeInfo.set_index('Number'), on='Number', how='left', rsuffix='_employee')

		details = detailData.loc[detailData['CLIN'] == clin].copy()
		
		details['AmountReg'] = details['HoursReg'] * details['BillRateReg']
		details['AmountOT'] = details['HoursOT'] * details['BillRateOT']
		details['AmountTotal'] = details['AmountReg'] + details['AmountOT']
		details['Wages'] = details['HoursTotal'] * details['HourlyRateReg']
		details['PostBill'] = (details['Wages'] * details['Posting'])
		details['HazardBill'] = (details['Wages'] * details['Hazard'])

		details = details[[
			'Date', 'Location', 'City', 'SubCLIN', 'Category', 'Name', 
			'Regular', 'LocalHoliday', 'Holiday', 'Vacation', 'Admin', 'Bereavement', 
			'Overtime', 'On-callOT', 'ScheduledOT', 'UnscheduledOT', 
			'HoursReg', 'BillRateReg', 'AmountReg', 
			'HoursOT', 'BillRateOT', 'AmountOT', 
			'HoursTotal', 'AmountTotal', 
			'HourlyRateReg', 'Posting', 'PostBill', 'Hazard', 'HazardBill'
		]]

		details.sort_values(['Date', 'Name'], inplace=True)

		############################################################################################################
		# Create the Costs file
		############################################################################################################

		outputFile = f'Costs-{activityYear}{activityMonth}-{region}-v3.xlsx'

		costData = details.copy()

		with pd.ExcelWriter(outputFile) as writer:
			if verbose:
				print(f'\nCreating {outputFile}...')

			# prepare a dictionary of information for the formatting function
			sheetInfo = {}
			sheetName = f'Costs-{region}'

			posts = details.groupby(['Location', 'City'], as_index=False).agg({'PostBill': 'sum'})
			posts['CLIN'] = '207'
			posts['Location'] = np.where(posts['Location'] == posts['City'], posts['Location'], posts['City'] + ', ' + posts['Location'])
			posts['City'] = 'Post'
			posts['G&A'] = posts['PostBill'] * upchargeRate
			posts['Total'] = posts['PostBill'] + posts['G&A']
			posts.rename(columns={'City': 'Type', 'PostBill': 'Amount'}, inplace=True)
			posts = posts[['CLIN', 'Location', 'Type', 'Amount', 'G&A', 'Total']]

			hazards = details.groupby(['Location', 'City'], as_index=False).agg({'HazardBill': 'sum'})
			hazards['CLIN'] = '208'
			hazards['Location'] = np.where(hazards['Location'] == hazards['City'], hazards['Location'], hazards['City'] + ', ' + hazards['Location'])
			hazards['City'] = 'Hazard'
			hazards['G&A'] = hazards['HazardBill'] * upchargeRate
			hazards['Total'] = hazards['HazardBill'] + hazards['G&A']
			hazards.rename(columns={'City': 'Type', 'HazardBill': 'Amount'}, inplace=True)
			hazards = hazards[['CLIN', 'Location', 'Type', 'Amount', 'G&A', 'Total']]

			costs = pd.concat([posts, hazards])

			
			# remove any row that has a total of zero
			# costs.drop(costs[costs.Total == 0].index, inplace=True)
			costs.sort_values(['CLIN', 'Location', 'Type'], inplace=True)

			invoiceAmount = costs['Total'].sum()
			costs = costs.loc[costs['Total'] > 0]

			if invoiceAmount > 0:
				summaryStartRow = 22
				rowsToSum = []

				rows = costs.shape[0]
				costs.to_excel(writer, sheet_name=sheetName, startrow=summaryStartRow, startcol=0, header=False)
				rowsToSum.append((summaryStartRow + 1, summaryStartRow + rows))
				summaryStartRow = summaryStartRow + rows + 2

				# use the first character of the region to uniquely identify this invoice
				uniqueRegion = region[0].upper()

				sheetInfo[sheetName] = {
					'invoiceNumber': costInvoiceNumber + uniqueRegion,
					'taskOrder': f'ODC-{region}',
					'dateStart': billing.dateStart.strftime("%m/%d/%Y"),
					'dateEnd': billing.dateEnd.strftime("%m/%d/%Y"),
					'invoiceAmount': invoiceAmount,
					'rowsToSum': rowsToSum
				}

		# Apply formatting in place
		workbook = load_workbook(outputFile)

		for styleName in styles.keys():
			workbook.add_named_style(styles[styleName])

		key = f'Costs-{region}'

		try:
			worksheet = workbook[key]
			info = sheetInfo[key]
			formatCostsTab(worksheet, info)
			
		except KeyError:
			print(f'No costs for {region}')
			pass

		workbook.save(outputFile)
		print(f'Invoices available at {outputFile}')

		############################################################################################################
		# Create the details file
		############################################################################################################
		
		outputFile = f'Details-{activityYear}{activityMonth}-{region}-v3.xlsx'

		# There is only one tab in the workbook
		sheetName = f'Details-{region}'

		with pd.ExcelWriter(outputFile) as writer:
			if verbose:
				print(f'\nCreating {outputFile}...')

			details.to_excel(writer, sheet_name=sheetName, startrow=0, startcol=0, header=True)
			
		# Apply formatting in place
		workbook = load_workbook(outputFile)

		for styleName in styles.keys():
			workbook.add_named_style(styles[styleName])
		
		worksheet = workbook[sheetName]
		formatDetailTab(worksheet)
			
		workbook.save(outputFile)
		print(f'Details available at {outputFile}')