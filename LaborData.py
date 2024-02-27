#!/usr/local/bin/python
import pandas as pd

from EmployeeTime import EmployeeTime
from BillingRates import BillingRates

# TERMINOLOGY:
# CLIN: The identifier for the region (e.g. "CLIN 001")
# Location: A country location within a region (e.g. "Russia")
# City: The city of the embassy where the work is performed (e.g. "Moscow")

# class LaborDetail:
# 	def __init__ (self, identifier:str, description:str, name:str, hours:float, rate:float, amount:float):
# 		self.identifier = identifier
# 		self.description = description
# 		self.name = name
# 		self.hours = hours
# 		self.rate = rate
# 		self.amount = amount

# class PostDetail:
# 	def __init__ (self, identifier:str, name:str, hours:float, hourlyRate:float, wages:float, city:str, rate:float, amount:float):
# 		self.identifier = identifier
# 		self.name = name
# 		self.hours = hours
# 		self.hourlyRate = hourlyRate
# 		self.wages = wages
# 		self.city = city
# 		self.rate = rate
# 		self.amount = amount

# class HoursSummary:
# 	def __init__ (self, city:str, identifier:str, name:str, regular:float, localHoliday:float, admin:float, overtime:float, onCallOT:float, scheduledOT:float, unscheduledOT:float, subtotal: float):
# 		self.city = city
# 		self.identifier = identifier
# 		self.name = name
# 		self.regular = regular
# 		self.localHoliday = localHoliday
# 		self.admin = admin
# 		self.overtime = overtime
# 		self.onCallOT = onCallOT
# 		self.scheduledOT = scheduledOT
# 		self.unscheduledOT = unscheduledOT
# 		self.subtotal = subtotal

# class HoursDetail:
# 	def __init__ (self, date:str, name:str, identifier:str, regular:float, localHoliday:float, admin:float, overtime:float, onCallOT:float, scheduledOT:float, unscheduledOT:float, subtotal: float):
# 		self.date = date
# 		self.name = name
# 		self.identifier = identifier
# 		self.regular = regular
# 		self.localHoliday = localHoliday
# 		self.admin = admin
# 		self.overtime = overtime
# 		self.onCallOT = onCallOT
# 		self.scheduledOT = scheduledOT
# 		self.unscheduledOT = unscheduledOT
# 		self.subtotal = subtotal

class LocationDetail:
	def __init__(self, locationName:str):
		self.locationName = locationName
		self.laborDetails = []
		self.postDetails = []
		self.hazardDetails = []
		self.hoursSummary = []
		self.hoursDetail = []

	def addLaborDetails(self, dataframe: pd.DataFrame):
		self.laborDetails.append(dataframe)

	def addPostDetails(self, dataframe: pd.DataFrame):
		self.postDetails.append(dataframe)

	def addHazardDetails(self, dataframe: pd.DataFrame):
		self.hazardDetails.append(dataframe)

	def addHoursSummary(self, dataframe: pd.DataFrame):
		self.hoursSummary.append(dataframe)
	
	def addHoursDetail(self, dataframe: pd.DataFrame):
		self.hoursDetail.append(dataframe)	

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

	def addHoursSummary(self, locationName: str, dataframe: pd.DataFrame):
		location = self.retrieveLocation(locationName)
		location.addHoursSummary(dataframe)
		self.locationDetails[locationName] = location

	def addHoursDetail(self, locationName: str, dataframe: pd.DataFrame):
		location = self.retrieveLocation(locationName)
		location.addHoursDetail(dataframe)
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
				##########################################################################
				# Data for labor invoices
				##########################################################################
				laborData = activity.groupedForInvoicing(clin=clin, location=location)

				for subCLIN in laborData['SubCLIN'].unique():
					clinData = laborData[laborData['SubCLIN'] == subCLIN]
					invoiceData.addLaborDetail(location, clinData)

				##########################################################################
				# Data for post invoices
				##########################################################################
				postData = activity.groupedForPostReport(clin=clin, location=location)

				# note: the "CLIN" column in postData is actually the subCLIN
				for subCLIN in postData['CLIN'].unique():
					clinData = postData[postData['CLIN'] == subCLIN]
					invoiceData.addPostDetail(location, clinData)

				##########################################################################
				# Data for hazard invoices
				##########################################################################
				hazardData = activity.groupedForHazardReport(clin=clin, location=location)

				# note: the "CLIN" column in postData is actually the subCLIN
				for subCLIN in hazardData['CLIN'].unique():
					clinData = hazardData[hazardData['CLIN'] == subCLIN]
					invoiceData.addHazardDetail(location, clinData)

				##########################################################################
				# Detail for hours report
				##########################################################################
				hoursSummary = activity.groupedForHoursReport(clin=clin, location=location)
				invoiceData.addHoursSummary(location, hoursSummary)

				hoursDetails = activity.byDate(clin=clin, location=location)
				invoiceData.addHoursDetail(location, hoursDetails)

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
			locationData = invoiceData.locationDetails[locationName]

			print(f'\nLabor Details: {len(locationData.laborDetails)}')
			for item in locationData.laborDetails:
				print(item)
			
			print(f'\nPost Details: {len(locationData.postDetails)}')
			for item in locationData.postDetails:
				print(item)

			if len(locationData.hazardDetails) > 0:
				print(f'\nHazard Details: {len(locationData.hazardDetails)}')
				for item in locationData.hazardDetails:
					print(item)

			print(f'\nHours Summary: {len(locationData.hoursSummary)}')
			for item in locationData.hoursSummary:
				print(item)
			
			print(f'\nHours Detail: {len(locationData.hoursDetail)}')
			for item in locationData.hoursDetail:
				print(item)

			print(f' ')