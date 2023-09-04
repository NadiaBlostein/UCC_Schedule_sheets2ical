'''
Quick script that allows user to download UCC graduate entry medicine (GEM) schedule into excel spreadsheet to convert it into icalendar event (instead of needing to do this manually).
Relevant documentation:
    * https://icalendar.readthedocs.io/en/latest/usage.html
    * https://learnpython.com/blog/working-with-icalendar-with-python/
'''

from icalendar import Calendar, Event
from datetime import datetime
import pandas as pd
import openpyxl

excel_file_name = 'GEM1 Semester 1 2023 2024.xlsx'

# Initiate the calendar
cal = Calendar()

# Some properties are required to be compliant
cal.add('prodid', '-//My calendar product//example.com//')
cal.add('version', '2.0')

# Function that creates dictionary out of specific worksheet
def get_ical(excel_file_name, worksheet):

    # Read dataframe
    try: df = pd.read_excel(excel_file_name, sheet_name = worksheet)
    except FileNotFoundError:
        print(f"Worksheet '{worksheet}' not found in the Excel file {excel_file}.")
        exit(1)
    df = df.reset_index(drop=True)
    df.columns = df.iloc[0]
    df = df[1:].reset_index(drop=True)
    df = df.drop(0).reset_index(drop=True)
    df.set_index(df.columns[0], inplace=True)
    df.fillna(0, inplace=True)
    
    # Dictionary with all the relevant information
    events = []
    for column_name, column_data in df.items():
        date = column_name #print(date.split()[0])
        for i, value in enumerate(column_data):
            if value != 0:
                event = {}
                event['summary'] = value

                dtstart = f"{date.year}-{date.month}-{date.day} {df.index.tolist()[i].split('-')[0].rstrip()}"
                dtstart = datetime.strptime(dtstart, "%Y-%m-%d %H:%M")
                event['dtstart'] = dtstart

                dtend = f"{date.year}-{date.month}-{date.day} {df.index.tolist()[i].split('-')[1].rstrip()}"
                dtend = datetime.strptime(dtend, "%Y-%m-%d %H:%M")
                event['dtend'] = dtend

                events.append(event)
    return(events)

sheet_list = openpyxl.load_workbook(excel_file_name).sheetnames

for i in range(5): # for worksheet in sheet_list:
    event_dict = get_ical(excel_file_name, worksheet = sheet_list[i])
    # Add subcomponents
    for course in event_dict:
        event = Event()
        event.add('summary', course['summary'])
        event.add('dtstart', course['dtstart'])
        event.add('dtend', course['dtend'])
        #event.add('location', course['Location'])
        cal.add_component(event)

f = open('GEM1_Semester_1_2023_2024.ics', 'wb')
f.write(cal.to_ical())
f.close()