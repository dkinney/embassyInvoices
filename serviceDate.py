#!/usr/local/bin/python
import sys
import pandas as pd
import calendar

from EmployeeTime import EmployeeTime


if len(sys.argv) < 1:
    print("Usage: python3 {sys.argv[0]} <data.csv>")
    sys.exit(1)

filename = sys.argv[1]
df = pd.read_csv(filename)
# df = df.loc[df['Project Name'].str.startswith('19AQMM23C0047')]
# print(df)

grouped = df.groupby(['Employee ID'], as_index=False).agg({
    'Employee Name': 'first', 
    'Entry Date': 'min'
})

grouped.to_csv('serviceDate.csv', index=False)