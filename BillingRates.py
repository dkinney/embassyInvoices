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
		self.effectiveDate = pd.to_datetime(effectiveDate) if effectiveDate is not None else pd.to_datetime('today')

		ratesFilename = filename if filename is not None else 'data/BillingRates.xlsx'

		print(f'Getting billing rates effective as of {effectiveDate} from {ratesFilename}')
		self.loadData(ratesFilename, verbose)

		employees = EmployeeInfo(verbose=verbose)
		self.joinWith(employees)

		if verbose:
			print(f'Loaded {len(self.data)} labor rates from {filename}')
			print(self.data.info())
			print(self.data)

	def loadData(self, filename, verbose=False):
		print('Loading billing rates from', filename)

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
		df = df.groupby('SubCLIN').apply(lambda x: x.loc[x['EffectiveDate'] <= self.effectiveDate].sort_values(by='EffectiveDate', ascending=True).head(1)).reset_index(drop=True)

		if verbose:
			print(f'Loaded {len(df)} labor rates from {filename}')
			print(df)

		self.data = df

	
	def joinWith(self, employees: EmployeeInfo):
		if employees.data is None:
			# nothing to do
			return

		joined = employees.data.join(self.data.set_index('SubCLIN'), on='SubCLIN', how='left', rsuffix='_rates')
		
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
	
	# Uses the file "BillingRates.xlsx" within the data directory

	if len(sys.argv) < 2:
		print(f'Usage: {sys.argv[0]} [YYYY-MM-DD]')
		sys.exit(1)

	# Uses the effective date provided, or today's date if not provided
	date = sys.argv[1] if len(sys.argv) > 1 else None

	# effectiveDate = pd.to_datetime(date)
	billingRates = BillingRates(effectiveDate=date, verbose=False)


	outputFile = 'EmployeeBillingRates.xlsx'

	tabData = {}

	billingRates.data.sort_values(by=['CLIN', 'SubCLIN', 'EmployeeName'], ascending=True, inplace=True)
	billingRates.data.fillna('Unknown', inplace=True)

	print(f'{len(billingRates.data)} employees for clins {billingRates.data["CLIN"].unique()}')

	for clin in billingRates.data['CLIN'].unique():
		try:
			region = Regions[clin]
		except KeyError:
			region = 'Unknown'

		data = billingRates.data[billingRates.data['CLIN'] == clin]
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
