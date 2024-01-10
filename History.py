import pandas as pd

class History:
	def __init__(self, filename=None, verbose=False):
		self.data = None	# a dataframe containing information loaded from a file and cleaned
		self.filename = filename if filename is not None else 'data/History.xlsx'
		self.loadData(self.filename, verbose)

		# if verbose:
		print(f'Loaded {len(self.data)} records from {self.filename}')
		print(self.data)
		
	def loadData(self, filename, verbose=False):
		# Define the data type will be used when reading in the data
		# By default, it will try to make columns that only have numbers into numbers.
		# converters = {
		# }
		df = pd.read_excel(filename, header=0)
		df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

		if verbose:
			print(f'Loaded {len(df)} records from {filename}')
			print(df)

		self.data = df

	def writeData(self, invoiceNumber, invoiceAmount):
		print(f'Writing {invoiceAmount} for {invoiceNumber} in {self.filename}')
		self.data.loc[len(self.data)] = [invoiceNumber, invoiceAmount]
		# self.data.drop_duplicates(keep='last', inplace=True)
	
		with pd.ExcelWriter(self.filename) as writer:
			self.data.to_excel(writer, sheet_name='History', startrow=0, startcol=0, index=False)

# The following is only used when testing this module.
# It expects a file in the parent directory for testing purposes only.

if __name__ == '__main__':
    import sys

    # By default, uses the file, "EmployeeInfo.xlsx" within the data directory
    # unless a filename is provided as a command line argument.
    inputFilename = sys.argv[1] if len(sys.argv) > 1 else None
    history = History(inputFilename, verbose=True)