import pandas as pd

class EmployeeInfo:
	def __init__(self, filename=None, verbose=False):
		self.data = None	# a dataframe containing information loaded from a file and cleaned

		filename = filename if filename is not None else 'data/EmployeeInfo.xlsx'

		# Define the data type will be used when reading in the data
		# By default, it will try to make columns that only have numbers into numbers.
		converters = {
			'EmployeeName': str,
			'EmployeeID': str,
			'HourlyRateReg': float,
			'HourlyRateOT': float,
			'Location': str,
			'Title': str,
			'PostingRate': float,
			'HazardRate': float
		}
	
		df = pd.read_excel(filename, header=0, converters=converters)
		df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

		self.checkMissing(df)

		# Ensure there are no duplicate rows
		# TODO: Make this an error so that it can be highlighted to clean the input data.
		df = df.drop_duplicates()

		# drop any employee that does not have an EmployeeID
		df = df.dropna(axis=0, how='any', subset=['EmployeeID'])

		# strip whitespace from all string columns
		df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

		if verbose:
			print(f'Loaded {len(df)} employees from {filename}')
			print(df)

		self.data = df

	def checkMissing(self, df: pd.DataFrame):
		self.missingNumber = df.loc[df['EmployeeName'].notna() & df['EmployeeID'].isna()]
		self.missingSubCLIN = df.loc[df['EmployeeName'].notna()  & df['EmployeeID'].notna() & df['SubCLIN'].isna()]

		employeeSet = set()
		employeeSet.update(self.missingSubCLIN['EmployeeID'].tolist())
		self.employeesMissingData = list(employeeSet)

	def describeMissing(self):
		if len(self.missingNumber) > 0:
			print(f"  Missing Number: {len(self.missingNumber)}")
			print(self.missingNumber[['EmployeeName']])

		if len(self.missingSubCLIN) > 0:
			print(f"  Missing SubCLIN: {len(self.missingSubCLIN)}")
			print(self.missingSubCLIN[['EmployeeName', 'EmployeeID']])

	def joinWith(self, billingRates):
		if billingRates.data is None:
			# nothing to do
			return

		joined = self.data.join(billingRates.data.set_index('SubCLIN'), on='SubCLIN', how='left', rsuffix='_rates')
		self.data = joined

# The following is only used when testing this module.
# It expects a file in the parent directory for testing purposes only.

if __name__ == '__main__':
	import sys
	from BillingRates import BillingRates

	# By default, uses the file, "EmployeeInfo.xlsx" within the data directory
	# unless a filename is provided as a command line argument.
	inputFilename = sys.argv[1] if len(sys.argv) > 1 else None
	employees = EmployeeInfo(inputFilename, verbose=False)

	# load PostHazard rates
	# postHazard = pd.read_csv('data/PostHazardRates.csv')
	# postHazard['EffectiveDate'] = pd.to_datetime(postHazard['EffectiveDate'])
	# postHazard.sort_values(by='EffectiveDate', inplace=False)
	# print('\nPostHazard Rates:')
	# print(postHazard)

	# load BillingRates
	from BillingRates import BillingRates
	billingRates = BillingRates(verbose=False)

	employees.joinWith(billingRates)
	debug = employees.data.loc[employees.data['EmployeeID'] == '11956']
	print(debug)
	print(debug[['EmployeeName', 'EmployeeID', 'City', 'PostingRate', 'HazardRate']])

	# billingRates.joinWith(employees)

	# debug = billingRates.data.loc[billingRates.data['EmployeeID'] == '11956']
	# print(debug)

	# # reordering the columns
	# billingRates.data = billingRates.data[[
	# 	'EmployeeName', 'EmployeeID', 'EffectiveDate',
	# 	'Title', 'HourlyRateReg', 'HourlyRateOT', 
	# 	'Location', 'City', 'PostingRate', 'HazardRate', 
	# 	'CLIN', 'SubCLIN', 'Category', 'BillRateReg', 'BillRateOT'
	# ]]

	# outputFile = 'data/EmployeeInfo.csv'

	# # Save the data to a csv file
	# billingRates.data.to_csv(outputFile, index=False)

	
	# with pd.ExcelWriter(outputFile) as writer:
	# 	billingRates.data.to_excel(writer, sheet_name='Info', startrow=0, startcol=0, header=True, index=False)