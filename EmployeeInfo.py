import pandas as pd

class EmployeeInfo:
	def __init__(self, filename=None, verbose=False):
		self.data = None	# a dataframe containing information loaded from a file and cleaned

		if filename is None:
			self.loadData('data/EmployeeInfo.xlsx', verbose)
		else:
			self.loadData(filename, verbose)

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

	def loadData(self, filename, verbose=False):
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

# The following is only used when testing this module.
# It expects a file in the parent directory for testing purposes only.

if __name__ == '__main__':
	import sys

	# By default, uses the file, "EmployeeInfo.xlsx" within the data directory
	# unless a filename is provided as a command line argument.
	inputFilename = sys.argv[1] if len(sys.argv) > 1 else None
	employees = EmployeeInfo(inputFilename, verbose=False)

	from BillingRates import BillingRates
	billingRates = BillingRates(verbose=False)
	billingRates.joinWith(employees)

	# reordering the columns
	billingRates.data = billingRates.data[[
		'EmployeeName', 'EmployeeID', 'EffectiveDate',
		'Title', 'HourlyRateReg', 'HourlyRateOT', 
		'Location', 'City', 'PostingRate', 'HazardRate', 
		'CLIN', 'SubCLIN', 'Category', 'BillRateReg', 'BillRateOT'
	]]

	outputFile = 'Resolved Employee Info.xlsx'
	with pd.ExcelWriter(outputFile) as writer:
		billingRates.data.to_excel(writer, sheet_name='Info', startrow=0, startcol=0, header=True, index=False)

	from openpyxl import load_workbook
	from InvoiceStyles import styles
	from InvoiceFormat import formatEmployeeInfo

	workbook = load_workbook(outputFile)

	for styleName in styles.keys():
		workbook.add_named_style(styles[styleName])
		
	worksheet = workbook['Info']
	formatEmployeeInfo(worksheet)
	workbook.save(outputFile)