#!/usr/local/bin/python
import pandas as pd

from EmployeeTime import EmployeeTime
from BillingRates import BillingRates

class LaborDetail:
	def __init__ (self, identifier:str, description:str, name:str, hours:float, rate:float, amount:float):
		self.identifier = identifier
		self.description = description
		self.name = name
		self.hours = hours
		self.rate = rate
		self.amount = amount

class PostDetail:
	def __init__ (self, identifier:str, name:str, hours:float, hourlyRate:float, wages:float, city:str, rate:float, amount:float):
		self.identifier = identifier
		self.name = name
		self.hours = hours
		self.hourlyRate = hourlyRate
		self.wages = wages
		self.city = city
		self.rate = rate
		self.amount = amount

class LocationDetail:
	def __init__(self, locationName:str):
		self.locationName = locationName
		self.laborDetails = []
		self.postDetails = []
		self.hazardDetails = []

	def addLaborDetails(self, dataframe: pd.DataFrame):
		for index, row in dataframe.iterrows():
			self.laborDetails.append(LaborDetail(row['SubCLIN'], row['Description'], row['EmployeeName'], row['Hours'], row['Rate'], row['Amount']))

	def addPostDetails(self, dataframe: pd.DataFrame):
		for index, row in dataframe.iterrows():
			self.postDetails.append(PostDetail(row['CLIN'], row['Name'], row['Regular'], row['Rate'], row['Regular Wages'], row['City'], row['Post Rate'], row['Post']))
	
	def addHazardDetails(self, dataframe: pd.DataFrame):
		for index, row in dataframe.iterrows():
			self.hazardDetails.append(PostDetail(row['CLIN'], row['Name'], row['Regular'], row['Rate'], row['Regular Wages'], row['City'], row['Hazard Rate'], row['Hazard']))

class InvoiceData:
	def __init__(self, clin:str):
		self.clin = clin
		self.locationDetails = {}

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

	def addPostDetail(self, locationName: str, dataframe: pd.DataFrame):
		location = self.retrieveLocation(locationName)
		location.addPostDetails(dataframe)
		self.locationDetails[locationName] = location
	
	def addHazardDetail(self, locationName: str, dataframe: pd.DataFrame):
		location = self.retrieveLocation(locationName)
		location.addHazardDetails(dataframe)
		self.locationDetails[locationName] = location

class LaborData:
	def __init__(self, activity: EmployeeTime):
		# create the data for creating labor invoices
		# actual creation happens separately

		# data is stored in a dictionary with the CLIN as the key
		self.invoiceData = {}

		locationInfo = activity.locationsByCLIN()
		
		for clin in locationInfo.keys():
			invoiceData = None

			if clin not in self.invoiceData:
				invoiceData = InvoiceData(clin)
			else:
				invoiceData = self.invoiceData[clin]

			for location in locationInfo[clin]:
				laborData = activity.groupedForInvoicing(clin=clin, location=location)

				for subCLIN in laborData['SubCLIN'].unique():
					clinData = laborData[laborData['SubCLIN'] == subCLIN]
					invoiceData.addLaborDetail(location, clinData)

				postData = activity.groupedForPostReport(clin=clin, location=location)

				# note: the "CLIN" column in postData is actually the subCLIN
				for subCLIN in postData['CLIN'].unique():
					clinData = postData[postData['CLIN'] == subCLIN]
					invoiceData.addPostDetail(location, clinData)

				hazardData = activity.groupedForHazardReport(clin=clin, location=location)

				# note: the "CLIN" column in postData is actually the subCLIN
				for subCLIN in hazardData['CLIN'].unique():
					clinData = hazardData[hazardData['CLIN'] == subCLIN]
					invoiceData.addHazardDetail(location, clinData)

			self.invoiceData[clin] = invoiceData

if __name__ == '__main__':
	import sys

	if len(sys.argv) < 2:
		print(f'Usage: {sys.argv[0]} <billing activity file>')
		sys.exit(1)

	filename = sys.argv[1]

	activity = EmployeeTime(filename, verbose=False)
	effectiveDate = activity.dateEnd
	billingRates = BillingRates(effectiveDate=effectiveDate, verbose=False)
	activity.joinWith(billingRates)

	labor = LaborData(activity)

	print(f'\nTesting data structures:')

	for clin in sorted(labor.invoiceData.keys()):
		invoiceData = labor.invoiceData[clin]
		print(f'CLIN: {clin}')

		for locationName in sorted(invoiceData.locationDetails.keys()):
			locationDetails = invoiceData.locationDetails[locationName]
			print(f'  Location {locationName} ({len(locationDetails.laborDetails)})')
			print(f'  Post {locationName} ({len(locationDetails.postDetails)})')
			print(f'  Hazard {locationName} ({len(locationDetails.hazardDetails)})')
			print(f' ')