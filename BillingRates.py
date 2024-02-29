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

		date = effectiveDate if effectiveDate is not None else 'today'
		self.effectiveDate = pd.to_datetime(date)

		ratesFilename = filename if filename is not None else 'data/BillingRates.xlsx'

		print(f'Getting billing rates effective as of {self.effectiveDate} from {ratesFilename}')

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
	
		df = pd.read_excel(ratesFilename, header=0, converters=converters)

		# Set BillRateReg and BillRateOT values to zero if they are NaN
		df['BillRateReg'].fillna(0, inplace=True)
		df['BillRateOT'].fillna(0, inplace=True)

		# strip whitespace from all string columns
		df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
		df = df.groupby('SubCLIN').apply(lambda x: x.loc[x['EffectiveDate'] <= self.effectiveDate].sort_values(by='EffectiveDate', ascending=True).head(1)).reset_index(drop=True)

		if verbose:
			print(f'Loaded {len(df)} labor rates from {ratesFilename}')
			print(df)

		self.data = df

if __name__ == '__main__':
	import sys

	# Uses the effective date provided, or today's date if not provided
	date = sys.argv[1] if len(sys.argv) > 1 else None

	# effectiveDate = pd.to_datetime(date)
	billingRates = BillingRates(effectiveDate=date, verbose=False)
	print(billingRates.data)