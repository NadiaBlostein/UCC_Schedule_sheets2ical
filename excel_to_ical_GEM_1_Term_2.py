from icalendar import Calendar, Event
from datetime import datetime
import pandas as pd
import openpyxl
import pytz

# =========== Get dictionary with all of the events
def get_ical(excel_file_name, worksheet):

    # Read dataframe
    try: df = pd.read_excel(excel_file_name, sheet_name = worksheet)
    except FileNotFoundError:
        print(f"Worksheet '{worksheet}' not found in the Excel file {excel_file_name}.")
        exit(1)
    df = pd.read_excel(excel_file_name, sheet_name = worksheet)
    df.columns = df.iloc[1,:]
    df = df[1:].reset_index(drop=True)
    df.set_index(df.iloc[:,0], inplace=True)

    # Dictionary with all the relevant information
    events = []
    idx_list = df.index.tolist()
    #print(idx_list)
    for column_name, column_data in df.items():
        if str(column_name) != 'NaT':
            #print(column_name)
            date = column_name
            if isinstance(date, datetime):
                
                for i, value in enumerate(column_data):
                    if value != 0 and isinstance(value,str):
                        
                        if 'lunch' not in value.lower():
                            event = {}
                            
                            # Parse strings to make it cleaner
                            value = value.replace('BHSC ','BHSC-')
                            value = value.replace('GM 1002','')
                            value = value.replace('GM1002','')
                            value = value.replace('M1002','')
                            value = value.replace('GM 1010','GM1010')
                            value = value.replace('GM 1020','GM1020')
                            value = value.replace('ANATOMY','Anatomy')
                            value = value.replace('PATHOLOGY /MEDMICRO','Pathology')
                            value = value.replace('PATHOLOGGY','Pathology')
                            value = value.replace('PHARMACOLOGY','Pharmacology')
                            value = value.replace('PHYSIOLOGY','Physiology')
                            value_list = value.split()
                            
                            # Separate location
                            filtered_value_list = [word for word in value_list if 'BHSC_' not in word]
                            location = [word for word in value_list if 'BHSC' in word]
                            output_value = ' '.join(filtered_value_list)
                            
                            event['summary'] = output_value

                            aa = idx_list[i]
                            if aa != 'nan' or aa != '':
                                aa = aa.replace('.',':')
                                dtstart = datetime(date.year, date.month, date.day,
                                    int(aa.split('-')[0].rstrip().split(':')[0]),0,0,
                                    tzinfo=pytz.timezone("GMT"))
                                event['dtstart'] = dtstart
                            
                                dtend = datetime(date.year, date.month, date.day,
                                    int(aa.split('-')[1].rstrip().split(':')[0]),0,0,
                                    tzinfo=pytz.timezone("GMT"))
                                event['dtend'] = dtend

                            if len(location) == 1: event['location'] = location[0]
                            events.append(event)
    return(events)

# =========== File name and sheet list
file_name = 'GEM_1_Term_2'
excel_file_name = file_name + '.xlsx'
sheet_list = openpyxl.load_workbook(excel_file_name).sheetnames

# =========== Initiate main calendar
cal = Calendar()
# =========== Some properties are required to be compliant
cal.add('prodid', '-//My calendar product//example.com//')
cal.add('version', '2.0')

# =========== Initiate module-specific calendars
module_cal = {}
module_list = ['Anatomy', 'Pathology','Pharmacology','Physiology', 'GM1010', 'GM1020', 'Misc','GEM_1_Term_2']
module_list_short = ['Anatomy', 'Pathology','Pharmacology','Physiology', 'GM1010', 'GM1020']
for module in module_list:
    tmp_cal = Calendar()
    tmp_cal.add('prodid', '-//My calendar product//example.com//')
    tmp_cal.add('version', '2.0')
    module_cal[module] = tmp_cal

# =========== Populate calendar with events!
# print(sheet_list)
for i in range(len(sheet_list)):
    if sheet_list[i] != 'Sheet3':
        # print(f'Current sheet {sheet_list[i]}')
        event_dict = get_ical(excel_file_name, worksheet = sheet_list[i])

        for course in event_dict:
            event = Event()
            event.add('summary', course['summary'])
            event.add('dtstart', course['dtstart'])
            event.add('dtend', course['dtend'])
            if course.get('location') is not None:
                event.add('location', course['location'])
            cal.add_component(event)
            if not any(item in course['summary'] for item in module_list_short):
                event = Event()
                event.add('summary', course['summary'])
                event.add('dtstart', course['dtstart'])
                event.add('dtend', course['dtend'])
                if course.get('location') is not None:
                    event.add('location', course['location'])
                module_cal['Misc'].add_component(event)
                module_cal['GEM_1_Term_2'].add_component(event)

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
                    module_cal['GEM_1_Term_2'].add_component(event)

# =========== Write icalendar files
f = open(file_name + '.ics', 'wb')
f.write(cal.to_ical())
f.close()
for module in module_list:
    f = open(module + '.ics', 'wb')
    f.write(module_cal[module].to_ical())
    f.close()