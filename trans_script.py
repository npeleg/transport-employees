# Created by Peleg Neufeld
# Assumptions:
#	- Addresses are not longer 7000 characters
#       - The work address exists

import xlrd
from xlutils.copy import copy
import requests
import datetime

# Parameters:
API_KEY = 'XXXX'
COUNTRY = 'IL'
GMT = 3
LANGUAGE = 'iw'
INPUT_FILE = "./example.xlsx"
OUTPUT_FILE = "./output.xls"

# Locations of data:
WORK_STREET = 0, 0
WORK_CITY = 0, 1
EMPLOYEES_STARTING_ROW = 2
EMPLOYEE_STREET = 1
EMPLOYEE_CITY = 2
EMPLOYEE_GEOLOCATION = 3
DRIVING_TO_WORK = 4
BICYCLING_TO_WORK = 5
TRANSIT_TO_WORK = 6
WALKING_TO_WORK = 7
DRIVING_FROM_WORK = 8
BICYCLING_FROM_WORK = 9
TRANSIT_FROM_WORK = 10
WALKING_FROM_WORK = 11

_errors_dict = {"ZERO_RESULTS": 'Could not find a valid address for employee. Please fill in a correct and accurate address.',
               "PARTIAL_MATCH": 'Could not find a valid address for employee. Please fill in a correct and accurate address.',
               "ZERO_DIRECTIONS_RESULTS": 'Origin or destination not found. Please fill in an accurate and correct address.'}


def get_address(sheet, street_row_col, city_row_col):
    address = sheet.cell_value(*street_row_col) + ", " + sheet.cell_value(*city_row_col)
    return address.replace(" ", "+")


def write_duration(sheet, row, travel_mode, direction, duration):
    mode_and_direction = travel_mode + direction
    col = globals()[mode_and_direction.upper()]
    sheet.write(row, col, duration)


def is_error(parsed_response):
    return parsed_response['status'] != "OK"


def get_error(parsed_response):
    status = parsed_response['status']
    error_title = _errors_dict[status] if status in _errors_dict else status
    if 'error_message' in parsed_response:
        error_title += ": " + parsed_response['error_message']
    return error_title


def is_partial_match(results):
    return 'partial_match' in results and results['partial_match'] is True


def get_local_hour(hour):
    local_hour = hour + GMT
    if local_hour < 0:
        return 24 + local_hour
    if local_hour > 24:
        return local_hour - 24
    return local_hour


def get_geocode_result(geocode_response):
    parsed_response = geocode_response.json()
    if is_error(parsed_response):
        return get_error(parsed_response), False
    if len(parsed_response['results']) > 1 or is_partial_match(parsed_response['results'][0]):
        return _errors_dict['PARTIAL_MATCH'], False
    location = parsed_response['results'][0]['geometry']['location']
    return str(location['lat']) + " " + str(location['lng']), True


def get_directions_result(directions_response):
    parsed_response = directions_response.json()
    if is_error(parsed_response):
        return get_error(parsed_response)
    for waypoint in parsed_response['geocoded_waypoints']:
        if waypoint['geocoder_status'] != 'OK':
            return _errors_dict['ZERO_DIRECTIONS_RESULTS']
        if is_partial_match(waypoint):
            return _errors_dict['PARTIAL_MATCH']

    route = parsed_response['routes'][0]
    if is_partial_match(route):
        return _errors_dict['PARTIAL_MATCH']
    leg = route['legs'][0]
    if 'duration_in_traffic' in leg:
        return leg['duration_in_traffic']['value'] / 60  # duration of commute, including traffic, in minutes
    else:
        return leg['duration']['value'] / 60  # duration of commute in minutes


# Open Workbook
input_file = INPUT_FILE
rb = xlrd.open_workbook(input_file)
r_sheet = rb.sheet_by_index(0)
wb = copy(rb)
w_sheet = wb.get_sheet(0)

work_address = get_address(r_sheet, WORK_STREET, WORK_CITY)


# Set departure times
date = datetime.datetime.now() + datetime.timedelta(days=1)  # tomorrow
if (COUNTRY == 'IL' and (date.weekday() == 4 or date.weekday() == 5)) or date.weekday() == 5 or date.weekday() == 6:
    date += datetime.timedelta(days=2)

morning_time = int(date.replace(hour=get_local_hour(8), minute=0, second=0).timestamp())  # 8:00 AM
evening_time = int(date.replace(hour=get_local_hour(17), minute=0, second=0).timestamp())  # 5:00 PM


# Traversing the employee list
for i in range(EMPLOYEES_STARTING_ROW, r_sheet.nrows):
    employee_street = i, EMPLOYEE_STREET
    employee_city = i, EMPLOYEE_CITY
    employee_address = get_address(r_sheet, employee_street, employee_city)

    # Filling employee's home address geolocation
    response = requests.get(f'https://maps.googleapis.com/maps/api/geocode/json?'
                            f'address={employee_address}'
                            f'&components=country:{COUNTRY}'
                            f'&language={LANGUAGE}'
                            f'&key={API_KEY}')
    result, found_address = get_geocode_result(response)
    w_sheet.write(i, EMPLOYEE_GEOLOCATION, result)

    # Filling in employee's commute times
    if not found_address:
        continue
    for direction in ['_to_work', '_from_work']:
        if direction == '_to_work':
            origin = employee_address
            destination = work_address
            departure_time = morning_time
        else:
            origin = work_address
            destination = employee_address
            departure_time = evening_time

        for travel_mode in ['driving', 'bicycling', 'transit', 'walking']:
            response = requests.get(f'https://maps.googleapis.com/maps/api/directions/json?'
                                    f'origin={origin}'
                                    f'&destination={destination}'
                                    f'&departure_time={departure_time}'
                                    f'&mode={travel_mode}'
                                    f'&language={LANGUAGE}'
                                    f'&key={API_KEY}')
            result = get_directions_result(response)
            write_duration(w_sheet, i, travel_mode, direction, result)

wb.save()
