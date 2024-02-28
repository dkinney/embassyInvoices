#!/usr/local/bin/python
import sys
import pandas as pd

def formatName(text):
	# print(f'formatName({text})')
	tokens = text.split(' ')

	if len(tokens) < 2:
			# not a name
			return text

	lastName = tokens[0]
	firstName = tokens[1]

	middleInitial = ''

	# hacktastic - eat extra spaces
	if len(tokens) > 3:
			middleInitial = tokens[3]
	elif len(tokens) > 2:
			middleInitial = tokens[2]

	if middleInitial != '':
		return f'{lastName}, {firstName} {middleInitial}'

	return f'{lastName}, {firstName}'

def outputByRegion(data: pd.DataFrame):
    for region in data['Region'].unique():
        with open(f'{region}-notApproved.tsv', 'w') as f:
            regionData = data.loc[data['Region'] == region].copy()
            regionData.drop(columns=['Region'], inplace=True)
            f.write(f'Hours Not Approved\tTotal\t{regionData["Duration"].sum()}\n')

            for name in sorted(regionData['Employee Name'].unique()):
                dataForOutput = regionData.loc[regionData['Employee Name'] == name].drop(columns=['Employee Name'])
                f.write(f'\n{name}\tSubtotal\t{regionData.loc[regionData["Employee Name"] == name, "Duration"].sum()}\n')
                f.write(dataForOutput.to_csv(index=False, header=False, sep='\t'))

# Read a contract number from argv[1] and filter the data in argv[2] for that contract number
if len(sys.argv) < 3:
    print("Usage: python3 filter.py [contractNumber] <data.csv>")
    sys.exit(1)

contractNumber = sys.argv[1]
filename = sys.argv[2]

data = pd.read_csv(filename, encoding='latin1')

# filter the data for the contract number
data = data.loc[data['Project Name'].str.startswith(contractNumber)]

if data.empty:
    print(f"No data found for contract number {contractNumber}")
    sys.exit(1)

# fix up the employee name
data['Employee Name'] = data['Employee Name'].apply(formatName)

# trim the first character from the Employee ID
data['Employee ID'] = data['Employee ID'].str[1:]

# add some columns to the filtered data from the 'Project Name' column
try:
    data['Contract ID'] = data['Project Name'].str.split().str[0]
    data['Region'] = data['Project Name'].str.split().str[2]
except Exception as e:
    print(f"Error: {e}")
    # this is not fatal, so we can continue
    pass

data.sort_values(by=['Entry Date', 'Region', 'Employee Name', 'State'], inplace=True)
grouped = data.groupby(['Region', 'Entry Date', 'Employee Name', 'State']).agg({'Duration': 'sum'}).reset_index()
notApproved = grouped.loc[grouped['State'] != 'Approved']

outputByRegion(notApproved)