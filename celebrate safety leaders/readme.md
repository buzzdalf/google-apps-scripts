This script finds a list of rows based on dates in a spreadsheet.  It copies summary information from each row into a new tab, and also into a google doc file.  Then it exports the google doc to a pdf and emails the file to someone so they can publish the pdf. This script executes on a scheduled trigger.
Concepts implemented:
- Finding cells based on dates
- clearing, copying to doc file
- deleting files from drive
- exporting to pdf in a specified drive folder
- email with doc file attached

current issues with the script:  The pdf export uses data from 1 export iteration old, not current data.  I am currently running the script twice via different starting functions and only emailing after the second run.  I tried to just iterate the script twice via one run of the script, but even that didn't work.  I also tried delaying the export up to 30 seconds in case I was outpacing drive, but that didn't work either.
