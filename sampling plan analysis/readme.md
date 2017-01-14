This script builds OPEX sampling plans and data entry sheets.  It also performs the analysis of the OPEX studies, allowing the user to analyze their data as a COV or an MSE.  It runs based off a custom menu, pulling the data for the sampling plan and data entry sheets from a sppreadsheet.  The analysis is all run from a custom sidebar.

Concepts implemented:
- building spreadsheets based on cell names
- conerting cell data into data inputs for manipulation & calculation
- converting spreadsheet data into arrays for manipulation, then writing the arrays back into spreadsheet
- array usage greatly increases the speed of execution vs trying to do everythign strictly using the spreadsheets for data
- custom sidebar for analysis
- alert dialog while script runs to hold the user and prevent issues with execution
- custom graphs for dot frequency diagrams and controls charts
- control chart implementation, including XBar & R charts and also I-MR charts.
