import pandas as pd

from EmployeeInfo import EmployeeInfo

class BillingRates:
	def __init__(self, filename=None, verbose=False):
		self.data = None			# a dataframe containing information loaded from a file and cleaned
		
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

		joined = employees.data.join(self.data.set_index('SubCLIN'), on='SubCLIN', how='left', rsuffix='_rates')

		# reorder the columns to be more useful
		joined = joined[['EmployeeName', 'EmployeeID', 'CLIN', 'SubCLIN', 'Location', 'City', 'Category', 'Title', 'HourlyRateReg', 'HourlyRateOT', 'BillRateReg', 'BillRateOT', 'EffectiveDate', 'PostingRate', 'HazardRate']]
		
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
	billingRates = BillingRates(billingRatesFilename, verbose=True)

	employees = EmployeeInfo(verbose=False)

	if employees.data is not None:
		billingRates.joinWith(employees)

	# do some analysis
	print(f'\nMargins by Employee:')
	margins = billingRates.marginByEmployeeSubCLIN()
	print(margins)