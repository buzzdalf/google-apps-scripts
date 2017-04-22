These scripts run a deviation process where the user starts in a document and selects their deviation based on a custom menu.  From there, they launch a form and enter the detailed information.  That form data then generates an email and gets logged into a spreadsheet.  90 days later a notice of expiration is sent.

Concepts implemented:
- populating specific locations of a doc file based on a custom menu
- applying custom text formatting to specific text in a document
- finding/clearing ranges in a document based on searches for specific text within the body
- launching a google form based on custom menu item
- reding input on form submit and using the files submitted to populate spreadsheet and email variables
- email, with addresses coming from a spreadsheet, and variables from form submittal
- sending follow-up emails based on date criteria since original form submittal (counting days)

The sheet script is tied to this spreadsheet: https://docs.google.com/spreadsheets/d/1JsrifxQ8iii03lLHt0y9liabvjCkmdDaZzmj84i4OsQ
The doc script is tied to this doc: https://docs.google.com/document/d/1onw9dgsKiPdA6MfufLHJncEaHivq-9W3EWE4Ktuohi4