from icalendar import Calendar, Event
from datetime import datetime
import pandas as pd
import openpyxl
import pytz
import os
from datetime import timedelta

# =========== Get dictionary with all of the events
def get_ical(excel_file_name, worksheet):

    # Read dataframe
    try: df = pd.read_excel(excel_file_name, sheet_name = worksheet)
    except FileNotFoundError:
        print(f"Worksheet '{worksheet}' not found in the Excel file {excel_file}.")
        exit(1)
    df = df.iloc[:, :6]
    df = df.reset_index(drop=True)
    df.columns = df.iloc[0]
    df = df[1:].reset_index(drop=True)
    df = df.drop(0).reset_index(drop=True)
    df.set_index(df.columns[0], inplace=True)
    
    # Dictionary with all the relevant information
    events = []
    for column_name, column_data in df.items():
        date = column_name
        if isinstance(date, datetime):
            for i, value in enumerate(column_data):
                if value != 0 and isinstance(value,str):
                    if 'lunch' not in value.lower():
                        event = {}
                        
                        # daylight savings reversed March 31st
                        tzone = "Europe/Dublin" #"Europe/Paris" if date.month > 3 else "GMT"

                        # Parse strings to make it cleaner
                        for module in ['Lecture','Anatomy', 'Biochemistry','Pathology','Pharmacology','Physiology', 'Special Studies Modules', 'Hospital Experience']:
                            value = value.replace(module.upper(),module)
                        for werd in ['BHSC_', 'BHSC ']: value = value.replace(werd,'BHSC-')
                        for werd in ['WGB__','WGB_', 'WGB ']: value = value.replace(werd,'WGB-')
                        value = value.replace('GM2001 Flame Lab Session','Anatomy\nFlame Lab Session')
                        value = value.replace('GM2001','')
                        value = value.replace('Special Studies modules','Special Studies Modules')
                        value_list = value.split()
                        
                        # Separate location
                        filtered_value_list = [word for word in value_list if 'BHSC-' not in word and 'WGB-' not in word]
                        location = [word for word in value_list if 'BHSC' in word or 'WGB' in word]
                        output_value = ' '.join(filtered_value_list)
                        
                        event['summary'] = output_value

                        dtstart = datetime(date.year, date.month, date.day,
                            int(df.index.tolist()[i].split('-')[0].rstrip().split(':')[0]),0,0,
                            tzinfo=pytz.timezone(tzone))
                        event['dtstart'] = dtstart
                        
                        if 'Flame Lab' in value: dtend = dtstart + timedelta(hours=2)
                        else:
                            dtend = datetime(date.year, date.month, date.day,
                                int(df.index.tolist()[i].split('-')[1].rstrip().split(':')[0]),0,0,
                                tzinfo=pytz.timezone(tzone))
                        event['dtend'] = dtend
        
                        if len(location) == 1: event['location'] = location[0]
                        events.append(event)
    return(events)

# =========== File name and sheet list
file_name = 'GEM_2_Term_1'
excel_file_name = file_name + '.xlsx'
sheet_list = openpyxl.load_workbook(excel_file_name).sheetnames

# =========== Initiate main calendar
cal = Calendar()

# =========== Some properties are required to be compliant
cal.add('prodid', '-//My calendar product//example.com//')
cal.add('version', '2.0')

# =========== Initiate module-specific calendars
module_cal = {}
module_list = ['Anatomy', 'Biochemistry','Pathology','Pharmacology','Physiology', 'GM2013', 'GM2020','Special Studies Modules','Misc']

for module in module_list:
    tmp_cal = Calendar()
    tmp_cal.add('prodid', '-//My calendar product//example.com//')
    tmp_cal.add('version', '2.0')
    module_cal[module] = tmp_cal

# =========== Populate calendar with events!
for i in range(1, 14):
    event_dict = get_ical(excel_file_name, worksheet = sheet_list[i])

    for course in event_dict:
        event = Event()
        event.add('summary', course['summary'])
        event.add('dtstart', course['dtstart'])
        event.add('dtend', course['dtend'])
        if course.get('location') is not None:
            event.add('location', course['location'])
        cal.add_component(event)
        if not any(item in course['summary'] for item in module_list):
            event = Event()
            event.add('summary', course['summary'])
            event.add('dtstart', course['dtstart'])
            event.add('dtend', course['dtend'])
            if course.get('location') is not None:
                event.add('location', course['location'])
            module_cal['Misc'].add_component(event)

    for module in module_list_short:
        for course in event_dict:
            if module in course['summary']:
                event = Event()
                event.add('summary', course['summary'])
                event.add('dtstart', course['dtstart'])
                event.add('dtend', course['dtend'])
                if course.get('location') is not None:
                    event.add('location', course['location'])
                module_cal[module].add_component(event)

# =========== Write icalendar files
if not os.path.exists(file_name): os.makedirs(file_name)
f = open(file_name + '/' + file_name + '.ics', 'wb')
f.write(cal.to_ical())
f.close()
for module in module_list:
    f = open(file_name + '/' + module + '.ics', 'wb')
    f.write(module_cal[module].to_ical())
    f.close()