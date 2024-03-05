#!/usr/local/bin/python
import pandas as pd
import re
import glob

from EmployeeTime import EmployeeTime
from EmployeeInfo import EmployeeInfo
from BillingRates import BillingRates
from Allowances import Allowances

# TERMINOLOGY:
# CLIN: The identifier for the region (e.g. "CLIN 001")
# Location: A country location within a region (e.g. "Russia")
# City: The city of the embassy where the work is performed (e.g. "Moscow")

def getUniquifier(pattern, type=None, region=None, year=None, monthName=None):
	# look for previous instances of the status file
	patternValues = re.findall(r'(.*)-(.*)-(.*)-(.*)', pattern)
	patternType = patternValues[0][0]
	patterRegion = patternValues[0][1]
	patternYear = patternValues[0][2]
	patternMonthhName = patternValues[0][3]

	statusFiles = glob.glob(pattern + '*.xlsx')

	version = 0
	for file in statusFiles:
		for vals in re.findall(r'(.*)-(.*)-(.*)-(.*)-(\d+)', file):
			typeMatches = type is None or type is not None and vals[0] == type
			regionMatches = region is None or region is not None and vals[1] == region
			yearMatches = year is None or year is not None and vals[2] == year
			monthMatches = monthName is None or monthName is not None and vals[3] == monthName

			if typeMatches and regionMatches and yearMatches and monthMatches:
				thisVersion = int(vals[len(vals)-1])
				version = max(version, thisVersion)

	version += 1
	return version

class LocationDetail:
	def __init__(self, locationName:str):
		self.locationName = locationName
		self.laborDetails = []
		self.laborHours = 0
		self.laborAmount = 0

		self.hoursSummary = []
		self.hoursDetail = []

	def addLaborDetails(self, dataframe: pd.DataFrame):
		self.laborDetails.append(dataframe)
		self.laborHours += dataframe['Hours'].sum()
		self.laborAmount += dataframe['Amount'].sum()

	# def addPostDetails(self, dataframe: pd.DataFrame):
	# 	self.postDetails.append(dataframe)
	# 	self.postAmount += dataframe['Post'].sum()

	# def addHazardDetails(self, dataframe: pd.DataFrame):
	# 	self.hazardDetails.append(dataframe)
	# 	self.hazardAmount += dataframe['Hazard'].sum()

	def addHoursSummary(self, dataframe: pd.DataFrame):
		self.hoursSummary.append(dataframe)
	
	def addHoursDetail(self, dataframe: pd.DataFrame):
		self.hoursDetail.append(dataframe)	

class InvoiceData:
	def __init__(self, clin:str):
		self.clin = clin
		self.locationDetails = {}
		self.hours = 0
		self.amount = 0

		self.postDetails: pd.DataFrame = None
		self.hazardDetails: pd.DataFrame = None

	def retrieveLocation(self, locationName:str) -> LocationDetail:
		try:
			location = self.locationDetails[locationName]
		except KeyError:
			location = LocationDetail(locationName)
		
		return location

	def addLaborDetail(self, locationName: str, dataframe: pd.DataFrame):
		location = self.retrieveLocation(locationName)
		location.addLaborDetails(dataframe)
		self.locationDetails[locationName] = location

	def addPostDetail(self, dataframe: pd.DataFrame):
		self.postDetails = dataframe
	
	def addHazardDetail(self, dataframe: pd.DataFrame):
		self.hazardDetails = dataframe

	def addHoursSummary(self, locationName: str, dataframe: pd.DataFrame):
		location = self.retrieveLocation(locationName)
		location.addHoursSummary(dataframe)
		self.locationDetails[locationName] = location

	def addHoursDetail(self, locationName: str, dataframe: pd.DataFrame):
		location = self.retrieveLocation(locationName)
		location.addHoursDetail(dataframe)
		self.locationDetails[locationName] = location

class LaborData:
	def __init__(self, time: EmployeeTime):
		# create the data for creating labor invoices
		# actual creation happens separately

		self.time = time

		# data is stored in a dictionary with the CLIN as the key

		# data organized for presentation in a labor invoice
		self.invoiceData = {}

		locationInfo = time.locationsByCLIN()
		
		for clin in locationInfo.keys():
			invoiceData = None

			if clin not in self.invoiceData:
				invoiceData = InvoiceData(clin)
			else:
				invoiceData = self.invoiceData[clin]

			# build the invoiceDetails for each location
			for location in locationInfo[clin]:
				##########################################################################
				# Data for labor invoices
				##########################################################################
				laborData = time.groupedForInvoicing(clin=clin, location=location)

				for role in laborData['RoleID'].unique():
					# print(f'laborData {location} - {role}')
					clinData = laborData[laborData['RoleID'] == role]
					invoiceData.addLaborDetail(location, clinData)

				##########################################################################
				# Detail for hours report
				##########################################################################
				hoursSummary = time.groupedForHoursReport(clin=clin, location=location)
				# print(f'hoursSummary: {location}')
				# print(hoursSummary)
				invoiceData.addHoursSummary(location, hoursSummary)

				hoursDetails = time.byDate(clin=clin, location=location)
				invoiceData.addHoursDetail(location, hoursDetails)

			##########################################################################
			# Data for post invoices
			##########################################################################
			postData = time.groupedForPostReport(clin=clin)
			invoiceData.addPostDetail(postData)

			hazardData = time.groupedForHazardReport(clin=clin)
			invoiceData.addHazardDetail(hazardData)

			self.invoiceData[clin] = invoiceData

	@classmethod
	def fromReportFile(cls, filename: str) -> 'LaborData':
		print(f'Creating labor data from {filename}')
		time = EmployeeTime(filename=filename)
		effectiveDate = time.dateEnd
		allowances = Allowances(effectiveDate=effectiveDate)
		billingRates = BillingRates(effectiveDate=effectiveDate)
		billingRates.joinWith(allowances)
		employees = EmployeeInfo()
		employees.joinWith(billingRates)
		time.joinWith(employees)
		return cls(time)

if __name__ == '__main__':
	import sys

	if len(sys.argv) < 2:
		print(f'Usage: {sys.argv[0]} <billing activity file>')
		sys.exit(1)

	reportFile = sys.argv[1]
	labor = LaborData.fromReportFile(reportFile)

	print(f'\nTesting data structures:')

	clin = '002'
	country = 'Russia'

	hours = labor.time.data
	hours = hours.loc[hours['CLIN'] == clin]
	hours = hours.loc[hours['Country'] == country]
	print(f'hours {country}: {hours.PostName.unique()}')

	print(hours)

	print(f'\nPivot:')

	# keep the existing columns but pivot the TaskName into columns
	pivot = hours.pivot_table(index=[
		'Date', 'CLIN', 'Country', 'PostName', 
		'RoleID', 'Category', 'EmployeeName', 
		'Rate', 'HourlyRate', 'PostingRate', 'HazardRate'
	], columns=['TaskName'], values=['Hours'], aggfunc='sum', fill_value=0).reset_index()


	print(pivot)
	print(f'hours {country}: {hours.PostName.unique()}')


	hoursSummary = labor.time.groupedForHoursReport(clin=clin, location=country)

	print(f'hoursSummary: {country}\n{hoursSummary}')

	# print(f'hoursSummary: {country}, {hoursSummary.PostName.unique()}')
	# print(hoursSummary)

	exit()

	for clin in sorted(labor.invoiceData.keys()):
		invoiceData = labor.invoiceData[clin]

		print(f'CLIN: {clin}')
		print(f'Locations: {sorted(invoiceData.locationDetails.keys())}')

		for locationName in sorted(invoiceData.locationDetails.keys()):
			invoiceDetails = invoiceData.locationDetails[locationName]

			# data = labor.time.groupedForInvoicing(clin=clin, location=locationName)
			# print(f'Location: {locationName}')
			# print(data)
			# print('\n')

		# print('\n-------------------------')
		# print(f'Post Details for {clin}: {len(invoiceData.postDetails)}')
		# print(invoiceData.postDetails)
		# summary = invoiceData.postDetails.groupby(['PostName']).agg({'Post': 'sum'}).reset_index()
		# print(summary)
		# print('---')

		# print('\n-------------------------')
		# print(f'Hazard Details for {clin}: {len(invoiceData.hazardDetails)}')
		# print(invoiceData.hazardDetails)
		# summary = invoiceData.hazardDetails.groupby(['PostName']).agg({'Hazard': 'sum'}).reset_index()
		# print(summary)
		# print('---')