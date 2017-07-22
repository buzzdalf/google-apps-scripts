This script manages a form & resulting spreadsheet database.  Audit areas in the form are generated based on the areas in the spreadsheet areas lsiting.  As the area listings in the sporeadsheet are added/removed, they are automatically added/removed from the form.
Additionally, As audit data is entered and fed into the spreadsheet, the listing of people needing audited in form dropdowns are then updated based on the elapsed time since that person was last audited. 
This creates a form with multiple dynamic dropdowns based on criteria established in the spreadsheet.

Concepts implemented:
- populate forms dropdowns based on spreadsheet lists
- update form dropdowns based on criteria from the spreadsheet
- maintain a master database in the results spreadsheet, using that to update for dropdowns
- calculations based on dates & elapsed time
- email error logging to a logging spreadsheet

This script is tied to: https://docs.google.com/spreadsheets/d/1Q5wPq6cfALLFh2sJY667DjSU3xrUaQgBrtm1WCk_L2Y
