#!/usr/local/bin/python
import pandas as pd

from BillingActivity import BillingActivity

if __name__ == '__main__':
	import sys

	activityFilename = sys.argv[1] if len(sys.argv) > 1 else None

	if activityFilename is None:
		print(f'Usage: {sys.argv[0]} <billing activity file>')
		exit()

	activity = BillingActivity(activityFilename, verbose=True)
	data = activity.data[['Date', 'EmployeeName', 'TaskName', 'Hours']]

	compareToFilename = sys.argv[2] if len(sys.argv) > 2 else None

	if compareToFilename is not None:
		compareTo = BillingActivity(compareToFilename, verbose=False)
		compareData = compareTo.data[['Date', 'EmployeeName', 'TaskName', 'Hours']]

		both = pd.concat([data, compareData], ignore_index=True)
		both = both.drop_duplicates(keep=False)
		print(f'Unique rows in {activityFilename} compared to {compareToFilename}')
		print(both)
	else:
		print(data)