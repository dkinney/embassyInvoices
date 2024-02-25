import pandas as pd
from openpyxl import load_workbook
from InvoiceStyles import styles
from InvoiceFormat import formatEmployeeInfo

from EmployeeInfo import EmployeeInfo
from PostHazard import PostHazard

Regions = {
	'001': 'Asia',
	'002': 'Europe'
}

class BillingRates:
	def __init__(self, filename=None, effectiveDate=None, verbose=False):
		self.data = None			# a dataframe containing information loaded from a file and cleaned
		self.effectiveDate = effectiveDate if effectiveDate is not None else pd.to_datetime('today').strftime('%Y-%m-%d')
		if filename is None:
			self.loadData('data/BillingRates.xlsx', verbose)
		else:
			self.loadData(filename, verbose)

		employees = EmployeeInfo(verbose=verbose)
		self.joinWith(employees)

		if verbose:
			print(f'Loaded {len(self.data)} labor rates from {filename}')
			print(self.data.info())
			print(self.data)

	def loadData(self, filename, verbose=False):
		# Define the data type will be used when reading in the data
		# By default, it will try to make columns that only have numbers into numbers.
		converters = {
			'CLIN': str,
			'SubClin': str,
			'SubCLIN': str,
			'Category': str,
			'Location': str,
			'BillRateReg': float,
			'BillRateOT': float
		}
	
		df = pd.read_excel(filename, header=0, converters=converters)

		# Set BillRateReg and BillRateOT values to zero if they are NaN
		df['BillRateReg'].fillna(0, inplace=True)
		df['BillRateOT'].fillna(0, inplace=True)

		# strip whitespace from all string columns
		df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

		# this returns ALL rates, including effective past and future
		if verbose:
			print(f'Loaded {len(df)} labor rates from {filename}')
			print(df)

		self.data = df
	
	def joinWith(self, employees: EmployeeInfo):
		if employees.data is None:
			# nothing to do
			return

		# filter the billing rates based on the effective date
		billingData = self.data.groupby('SubCLIN').apply(lambda x: x.loc[x['EffectiveDate'] <= self.effectiveDate].sort_values(by='EffectiveDate', ascending=False).head(1)).reset_index(drop=True)

		joined = employees.data.join(billingData.set_index('SubCLIN'), on='SubCLIN', how='left', rsuffix='_rates')

		# reorder the columns to be more useful
		joined = joined[['EmployeeName', 'EmployeeID', 'CLIN', 'SubCLIN', 'Location', 'City', 'Category', 'Title', 'HourlyRateReg', 'HourlyRateOT', 'BillRateReg', 'BillRateOT', 'EffectiveDate']]
		
		postHazard = PostHazard(effectiveDate=self.effectiveDate)
		joined = joined.join(postHazard.data.set_index('PostName'), on='City', how='left', rsuffix='_post')

		# reorder the columns to be more useful
		joined = joined[['EmployeeName', 'EmployeeID', 'EffectiveDate', 'Title', 'HourlyRateReg', 'HourlyRateOT', 'Location', 'City', 'PostingRate', 'HazardRate', 'CLIN', 'SubCLIN', 'Category', 'BillRateReg', 'BillRateOT']]

		# this overwrites the data attribute with the more detailed information
		self.data = joined
	
	def marginByEmployeeSubCLIN(self):
		# Produce a dataframe with the margin for each SubCLIN for analysis
		margins = self.data.copy()
		margins['MarginReg'] = margins['BillRateReg'] - margins['HourlyRateReg']
		margins['MarginRate'] = margins['MarginReg'] / margins['HourlyRateReg']
		margins.sort_values(by=['MarginRate', 'SubCLIN'], ascending=False, inplace=True)
		return margins[['MarginRate', 'MarginReg', 'SubCLIN', 'HourlyRateReg', 'BillRateReg', 'EmployeeName', 'Category']]

if __name__ == '__main__':
	import sys
	
	# By default, uses the file "BillingRates.xlsx" within the data directory
	# unless a filename is provided as a command line argument.
	billingRatesFilename = sys.argv[1] if len(sys.argv) > 1 else None

	effectiveDate = pd.to_datetime('2024-02-29').strftime('%Y-%m-%d')
	billingRates = BillingRates(billingRatesFilename, effectiveDate=effectiveDate, verbose=False)

	employees = EmployeeInfo(verbose=False)

	if employees.data is not None:
		billingRates.joinWith(employees)

	outputFile = 'EmployeeBillingRates.xlsx'

	tabData = {}

	billingRates.data['CLIN'].fillna('Unknown', inplace=True)
	print(f'{len(billingRates.data)} employees for {billingRates.data["CLIN"].unique()}')

	for clin in billingRates.data['CLIN'].unique():
		try:
			region = Regions[clin]
		except KeyError:
			region = 'Unknown'

		data = billingRates.data[billingRates.data['CLIN'] == clin]

		print(f'Writing to {outputFile} to sheet {region} for {len(data)} employees')
		print(data)

		tabData[region] = data

	with pd.ExcelWriter(outputFile) as writer:
		for tab in tabData.keys():
			data = tabData[tab]
			data.to_excel(writer, sheet_name=tab, startrow=0, startcol=0, header=True, index=False)

	# Apply formatting in place
	workbook = load_workbook(outputFile)

	for styleName in styles.keys():
		workbook.add_named_style(styles[styleName])
	
	for sheet in workbook.sheetnames:
		worksheet = workbook[sheet]
		formatEmployeeInfo(worksheet)
	
	workbook.save(outputFile)
