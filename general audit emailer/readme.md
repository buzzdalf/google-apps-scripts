This script sends emails out based on new form submittals.  It then updates the spreadsheet with a flag that emails were successfully sent. This code is based on the code written for walking audit emails, but is built to be applied more generally.

Concepts implemented:
- Finding columns by name, cleaning up column names by removing special characters, case changes, etc
- copying formulas from a range to another range of cells
- email, with addresses, subject, and body coming from a spreadsheet
- html emails with character formatting like text color changes and bold
- falling back to standard text if no html code is found in messge body cell in the spreadsheet
- email error logging to a logging spreadsheet

This script is tied to this spreadsheet: https://docs.google.com/spreadsheets/d/1Oh4z8oHqlpbokIz2HftcLMAzK6_yFJmWbL3jBbu2Byk