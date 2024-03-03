#!/usr/local/bin/python
import sys
import pandas as pd

from BillingRates import BillingRates
billingRates = BillingRates()

if len(sys.argv) < 2:
    print(f"Usage: python3 {sys.argv[0]} <data.csv>")
    sys.exit(1)

filename = sys.argv[1]
data = pd.read_excel(filename, header=3)

print(f'\nLoaded {len(data)} rows from {filename}')
# print(data[['P_CLIN', 'RT', 'OT']])
newData = data[['P_CLIN', 'RT', 'OT']]

# rename columns for internal usage
newData.rename(columns={
    'P_CLIN': 'P_CLIN',
    'RT': 'JoeRateReg',
    'OT': 'JoeRateOT'
}, inplace=True)

joined = billingRates.data.join(newData.set_index('P_CLIN'), on='RoleID', how='left', rsuffix='_joe')
print(joined.info())
# print(joined[['RoleID', 'BillRateReg', 'BillRateOT', 'RT', 'OT']])

joined['BillRateReg'] = joined['BillRateReg'].astype(float)
joined['BillRateOT'] = joined['BillRateOT'].astype(float)
joined = joined.fillna(0)

joined['DiffReg'] = joined['BillRateReg'] - joined['JoeRateReg']
joined['DiffOT'] = joined['BillRateOT'] - joined['JoeRateOT']

output = joined[['RoleID', 'BillRateReg', 'BillRateOT', 'JoeRateReg', 'JoeRateOT', 'DiffReg', 'DiffOT']]
print(output)
output.to_csv('billingRatesComparison.csv', index=False)
