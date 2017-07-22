This script scrapes a folder and all sub-folders, looking for files that meet criteria.  When one of the targeted files is found, the script then searches inside that file for contents of specific cells and uses that as another criteria.  If that criteria is met, data is extracted from the file and pasted into a "tracking" spreadsheet.  The tracking spreadsheet data is then copied into tabs based on different labels.  The tabs are each then formatted.
Concepts implemented:
- Searching google drive folders and all sub-folders for files based on criteria
- file modify date check against criteria
- find cells based on labels or titles
- Inspect spreadsheet contents for sign-offs by title
- format cells based on criteria
- email, with addresses, subject, and body content coming from a spreadsheet
- email error logging to a logging spreadsheet

This script is tied to this spreadsheet: https://docs.google.com/spreadsheets/d/13_3VwQDfvVM53kDgTj8pxXd3jDwlCzaUAjURAxvQP2k