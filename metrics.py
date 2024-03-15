#!/usr/local/bin/python
import sys
import pandas as pd

from EmployeeTime import EmployeeTime
from EmployeeInfo import EmployeeInfo
from BillingRates import BillingRates
from Allowances import Allowances

billingRates = BillingRates()
allowances = Allowances()
billingRates.joinWith(allowances)
employees = EmployeeInfo()
employees.joinWith(billingRates)

comparison = employees.data[['EmployeeName', 'EmployeeID', 'Country', 'PostName', 'RoleID', 'Title', 'Category']]
print(comparison)
comparison.to_csv('employeeInfoComparison.csv', index=False)

exit()

if len(sys.argv) < 2:
    print(f'Usage: {sys.argv[0]} <billing activity file>')
    sys.exit(1)

filename = sys.argv[1]
time = EmployeeTime(filename)
time.joinWith(employees)

print('\nEmployee Time:')
print(time.data)