# Packages
import datetime
import requests
from pprint import pprint

from openpyxl import Workbook, load_workbook
from darksky import forecast

# Constants
filename = "record.xlsx"
api_key = "51e2256fa5b19e8e9ac16c7bde4acc8a"

latitude = 33.7591
longitude = -118.3872

# Load Spreadsheet
workbook = load_workbook(filename)
worksheet = workbook.active

# Read Field Headers from row 2
fields = {}
for cell in worksheet['2']:
	if (type(cell.value) == str):
		fields[cell.col_idx - 1] = cell.value
	
# Find Last Date
for cell in reversed(list(worksheet['A'])): # iterate backwards
	date = cell.value
	if (type(date) == datetime.datetime):
		break
date = datetime.datetime.combine(date.date(), datetime.datetime.min.time()) # zero out the time

# Collect DarkSky Data
today = datetime.date.today()

def areSameDay(a, b):
	return a.day == b.day and a.month == b.month and a.year == b.year

date += datetime.timedelta(days = 1)
while not areSameDay(date, today): # foreach day
	
	weather = forecast(api_key, latitude, longitude, time = date.isoformat())	
	data_point = vars(weather.daily.data[0])
	
	for hourly_data in weather.hourly: # foreach hour
		data_point.update(vars(hourly_data))
		data_point["date"] = datetime.datetime.fromtimestamp(data_point["time"])
		data_point["hour"] = data_point["date"].hour
		
		if not areSameDay(data_point["date"], date):
			break
		
		new_row = [None] * (max(fields.keys()) + 1)
		for column, field in fields.items():
			new_row[column] = data_point[field]
		
		worksheet.append(new_row)
	
	date += datetime.timedelta(days = 1)
	
# Save Spreadsheet
workbook.save(filename)

print("DarkSky Parser: Success!")