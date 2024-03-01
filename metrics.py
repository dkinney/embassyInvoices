#!/usr/local/bin/python
import sys
import pandas as pd

from EmployeeTime import EmployeeTime
from EmployeeInfo import EmployeeInfo
from BillingRates import BillingRates
from Allowances import Allowances

if len(sys.argv) < 2:
    print(f'Usage: {sys.argv[0]} <billing activity file>')
    sys.exit(1)

filename = sys.argv[1]

time = EmployeeTime(filename)
billingRates = BillingRates()

allowances = Allowances()
billingRates.joinWith(allowances)

print('\nBilling Rates:')
print(billingRates.data)

employees = EmployeeInfo()

employees.joinWith(billingRates)
time.joinWith(employees)

print('\nEmployee Time:')
print(time.data)