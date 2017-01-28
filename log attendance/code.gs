/* Script to allow people to register their meeting attendance in a spreadsheet
*  It also grabs meeting information and loads it into the sheet for the meeting instance
*  This script is tied to: https://docs.google.com/spreadsheets/d/1HOrb3c-ZcSrjRvpOWRAr7YnVSEyHIR-DRCO8WUWOxOU
*  By: Bill Steinberger
*/
var ss = SpreadsheetApp.getActiveSpreadsheet();

function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('Log Attendance')
  .addItem('I Attended Today', 'buildSheets')
  .addToUi();
}

function buildSheets() {
  var sheets = ss.getSheets();
  var now = new Date();
  var date = (now.getMonth()+1) + "/" + now.getDate() + "/" +  now.getFullYear();
  var attendSheet = "attendance "+date;
  var exist = false;
  
  for (var i=0;i<sheets.length;i++) {
    var name = sheets[i].getName();
    if (name == attendSheet) {
      exist = true;
    }
  }
  
  if (!exist) {
    configureSheet(attendSheet);
  }
  getEmail(attendSheet,now);
}

function getEmail(mySheet,date) {
  var sheet= ss.getSheetByName(mySheet);
  ss.setActiveSheet(sheet);
  var Avals = ss.getRange("A1:A").getValues();
  var Alast = Avals.filter(String).length;
  var email = Session.getActiveUser().getEmail(); 
  var emptyRow = Alast + 1;
  
  sheet.getRange(emptyRow,1).setValue(date);
  sheet.getRange(emptyRow,2).setValue(email);
}

function configureSheet(addSheet) {
  ss.insertSheet(addSheet, 1);
  var sheet = SpreadsheetApp.getActiveSheet();
  
  getProForma(sheet);
  
  sheet.getRange(1,1).setValue('Date');
  sheet.getRange(1,2).setValue('Name');
  
}

function getProForma(outSheet) {
  var inSheet = ss.getSheetByName('Weekly Pro Forma for Meeting');
  var inData = inSheet.getRange(1,1,50,15).getValues();
  
  outSheet.getRange(1,1,50,15).setValues(inData);
}

