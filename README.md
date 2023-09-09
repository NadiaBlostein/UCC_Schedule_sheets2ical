# What 
* Quick script that allows user to download UCC graduate entry medicine (GEM) schedule into excel spreadsheet to convert it into icalendar event (instead of doing this manually).
* `GEM1_Semester_1_2023-2024_calendars` contains icalendar versions of every GEM1 Fall 2023 module, as well as a calendar containing our entire schedule.
* You may want to run the script used to automatically convert excel schedule to ical yourself (e.g. you are using a different Excel spreadsheet becuase you come from a different year, want a more up-to-date schedule, etc). Below is an explanation of how to do this.

# How?
### Prerequisites
* Familiarity with command-line and basic Python / pip wheel
* `Python 3.8.13`
* `pip 22.3.1`

### What to do
1. Download UCC course schedule as Excel file (`GEM1 Semester 1 2023 2024.xlsx`) and move it to local directory.
2. Run the following commands from your terminal
```
pip install -r requirements.txt
python excel_to_ical.py
```

### Caution
When you click on the output `ics` file, make sure to load it into its own new Calendar (when you update it, it will generate every event from scratch so you will want to delete the previous version).

# Pending features
* ability to specify the week (i.e. worksheet) and spreadsheet of interest from command line
* real-time updates