from icalendar import Calendar, Event, vCalAddress, vText
from datetime import datetime
from pathlib import Path
import os
import pytz
import pandas as pd
import openpyxl
import numpy as np

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
                        
                        # Parse strings to make it cleaner
                        value = value.replace('BHSC ','BHSC_')
                        value = value.replace('GM1001','')
                        value = value.replace('PHARMACOLOGY','Pharmacology')
                        value = value.replace('BIOCHEMISTRY','Biochemistry')
                        value = value.replace('ANATOMY','Anatomy')
                        value = value.replace('PATHOLOGY /MEDMICRO','Pathology/Micro')
                        value_list = value.split()
                        
                        # Separate location
                        filtered_value_list = [word for word in value_list if 'BHSC_' not in word]
                        location = [word for word in value_list if 'BHSC_' in word]
                        output_value = ' '.join(filtered_value_list)
                        
                        event['summary'] = output_value
                        dtstart = f"{date.year}-{date.month}-{date.day} {df.index.tolist()[i].split('-')[0].rstrip()}"
                        dtstart = datetime.strptime(dtstart, "%Y-%m-%d %H:%M")
                        event['dtstart'] = dtstart
                        dtend = f"{date.year}-{date.month}-{date.day} {df.index.tolist()[i].split('-')[1].rstrip()}"
                        dtend = datetime.strptime(dtend, "%Y-%m-%d %H:%M")
                        event['dtend'] = dtend
                        if len(location) == 1: event['location'] = location[0]
                        events.append(event)
    return(events)

# =========== File name and sheet list
file_name = 'GEM1 Semester 1 2023 2024'
excel_file_name = file_name + '.xlsx'
sheet_list = openpyxl.load_workbook(excel_file_name).sheetnames

# =========== Initiate the calendar
cal = Calendar()

# =========== Some properties are required to be compliant
cal.add('prodid', '-//My calendar product//example.com//')
cal.add('version', '2.0')

# =========== Populate calendar with events!
for i in range(14):
    event_dict = get_ical(excel_file_name, worksheet = sheet_list[i])

    for course in event_dict:
        event = Event()
        event.add('summary', course['summary'])
        event.add('dtstart', course['dtstart'])
        event.add('dtend', course['dtend'])
        if course.get('location') is not None:
            event.add('location', course['location'])
        cal.add_component(event)

# =========== Write icalendar file
f = open(file_name + '.ics', 'wb')
f.write(cal.to_ical())
f.close()
