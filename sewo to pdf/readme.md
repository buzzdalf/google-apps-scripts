This script finds files in a specific folder based on search criteria (date, file type, file name), copies the contents of a sheet in each file out to a temp spreadsheet, strips out un-wanted information, converts that to a pdf, and sends an email to people who have their name on the email list.   This script is still in development.  The emails are not yet turned on.
Concepts implemented:
- Searching google drive folders for files based on criteria
- file modify date check against criteria
- export sheet to a new spreadsheet
- modify spreadsheet content
- convert spreadsheet to pdf
- email, with addresses, subject, and body content coming from a spreadsheet
- email error logging to a logging spreadsheet

This script is tied to this spreadsheet: https://docs.google.com/spreadsheets/d/1Z9jJ0pVl7XvlkGe-ohWR55j-lE0jhMcdexFhzAPYDgg