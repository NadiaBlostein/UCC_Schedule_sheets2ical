# Prerequisites
```
Python 3.8.3
pip
```

# What to do
1. Download UCC course schedule as Excel file (`GEM1 Semester 1 2023 2024.xlsx`) and move it to local directory.
2. Run the following commands from your terminal
```
pip install -r requirements.txt
python excel_to_ical.py
```

# Caution
When you click on the output `ics` file, make sure to load it into its own new Calendar (when you update it, it will generate every event from scratch so you will want to delete the previous version).

# Pending features
* going beyond week 9 (bugs here – probably related to a typo)
* ability to specify the week (i.e. worksheet) of interest from command line
* parse location
* add module name to first line of event title
* ability to overwrite previously generated events (as opposed to writing an updated duplicate)
