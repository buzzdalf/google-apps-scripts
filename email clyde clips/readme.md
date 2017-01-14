This script emails out the clyde clips paper for the current week.  It runs on a schedule trigger daily, looking for a file in the specified folder containing today's date.  If it find's one, it attaches it to an email and sends it out.

Concepts implemented:
- compile email addresses based on cell contents
- Finding files in drive based on dates
- email with doc file attached
