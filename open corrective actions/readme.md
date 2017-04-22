This script will go through all standard safety audits and find open corrective actions
Then it checks to see if they are older than 30 days.  If they are, it adds them to the report
You can run the report generator from a custom menu item.
You can also set it to run automatically and email out the results on a schedule (set trigger to autoRun())

Concepts implemented:
- send emails based on matching names in one sheet to emails addresses in another sheet
- stack multiple items onto an output sheet from a row based on a key value
- Check date based on cell contents
- sorts ranges
- clears ranges
- email error logging to a logging spreadsheet

This script is tied to: https://docs.google.com/a/whirlpool.com/spreadsheets/d/1R-D7gBguUtWntgli6aq-CSYnWYZ_yXYopcHnnjzk4ds
