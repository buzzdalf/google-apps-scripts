/*
** This script will go through all standard safety audits and find open corrective actions
** Then it checks to see if they are older than 30 days.  If they are, it adds them to the report
** You can run the report generator from a custom menu item.
** You can also set it to run automatically and email out the results on a schedule (set trigger to autoRun())
** This script is tied to: https://docs.google.com/a/whirlpool.com/spreadsheets/d/1R-D7gBguUtWntgli6aq-CSYnWYZ_yXYopcHnnjzk4ds
** By: Bill Steinberger
*/

var ss = SpreadsheetApp.getActiveSpreadsheet();
var outputSheet = ss.getSheetByName('Open');
var pastdueSheet = ss.getSheetByName('Past Due');
var emailSheet = ss.getSheetByName('Email addresses');
var emailFlag = false;

function onOpen() {
SpreadsheetApp.getUi()
.createMenu('Refresh List')
.addItem('Update Open Audits', 'findSheets')
.addToUi();
}

// execute this function to run everything, including sending the emails
function autoRun() {
  emailFlag = true;
  findSheets();
  sendEmail();
}

// executing this function will update the spreadsheets, but not send any emails
function findSheets() {
  clearSheet(outputSheet);
  clearSheet(pastdueSheet);
  
  var environmentSheet = ss.getSheetByName('Environment');
  var equipmentSheet = ss.getSheetByName('Equipment');
  var materialsSheet = ss.getSheetByName('Materials');
  var walkingSheet = ss.getSheetByName('Walking-Working Surfaces');

  var environmentData = environmentSheet.getDataRange().getValues();
  var environmentOpen = findOpen(environmentData,'Environment');

  var equipmentData = equipmentSheet.getDataRange().getValues();
  var equipmentOpen = findOpen(equipmentData,'Equipment');

  var materialsData = materialsSheet.getDataRange().getValues();
  var materialsOpen = findOpen(materialsData,'Materials');
  
  var walkingData = walkingSheet.getDataRange().getValues();
  var walkingOpen = findOpen(walkingData,'Walking-Working Surfaces');
  
  sortSheet(outputSheet);
  sortSheet(pastdueSheet);
}


// this function is the core functionality, finding the audits and retrieving the cell data to populate the output sheets
function findOpen(data,audit) {
  var anyopenCol = findColumn('Number of Corrective Actions Open',data);  
  var monthCol = findColumn('Month', data);
  var lineCountCol = findColumn('Line Count',data);
  var areaCol = findColumn('Area',data);
  var emailCol = findColumn('Email Address',data);
  var dateCol = findColumn('Date & Time',data);
  var shiftCol = findColumn('Shift',data);
  var observerCol = findColumn('Observer Name',data);
  var buCol = findColumn('Business Unit',data);
  var ca1DescCol = findColumn('Corrective Action 1 Description',data) || findColumn('Corrective Action 1 Decription',data);
  var ca1RespCol = findColumn('Corrective Action 1 Responsible Party',data);
  var ca1DueCol = findColumn('Corrective Action 1 Due Date',data);
  var ca1CompCol = findColumn('Corrective Action 1 Completion Date',data);
  var ca2DescCol = findColumn('Corrective Action 2 Description',data) || findColumn('Corrective Action 2 Decription',data);
  var ca2RespCol = findColumn('Corrective Action 2 Responsible Party',data);
  var ca2DueCol = findColumn('Corrective Action 2 Due Date',data);
  var ca2CompCol = findColumn('Corrective Action 2 Completion Date',data);
  var ca3DescCol = findColumn('Corrective Action 3 Description',data);
  var ca3RespCol = findColumn('Corrective Action 3 Responsible Party',data);
  var ca3DueCol = findColumn('Corrective Action 3 Due Date',data);
  var ca3CompCol = findColumn('Corrective Action 3 Completion Date',data);
  var ca4DescCol = findColumn('Corrective Action 4 Description',data);
  var ca4RespCol = findColumn('Corrective Action 4 Responsible Party',data);
  var ca4DueCol = findColumn('Corrective Action 4 Due Date',data);
  var ca4CompCol = findColumn('Corrective Action 4 Completion Date',data);
  
 
  for (var i=1;i<data.length;i++) {
    var anyOpen = data[i][anyopenCol];
    var old = checkDate(data[i][dateCol]);
    var output = [[]];

    if (anyOpen > 0 && old) {
      var openCount = 0;
      var isOneOpen = checkOpen(data[i][ca1DueCol],data[i][ca1CompCol]);
      if (isOneOpen) {
        openCount++;
      }
      var isTwoOpen = checkOpen(data[i][ca2DueCol],data[i][ca2CompCol]);
      if (isTwoOpen) {
        openCount++;
      }
      var isThreeOpen = checkOpen(data[i][ca3DueCol],data[i][ca3CompCol]);
      if (isThreeOpen) {
        openCount++;
      }
      var isFourOpen = checkOpen(data[i][ca4DueCol],data[i][ca4CompCol]);
      if (isFourOpen) {
        openCount++;
      }
      
      for (var j=0;j<openCount;j++) { 
        output[0][0] = data[i][monthCol]; 
        output[0][1] = audit;
        output[0][2] = data[i][lineCountCol];
        output[0][3] = data[i][areaCol];
        output[0][4] = data[i][emailCol];
        output[0][5] = data[i][dateCol];
        output[0][6] = data[i][shiftCol];
        output[0][7] = data[i][observerCol];
        output[0][8] = data[i][buCol];
       
        var temp9, temp10, temp11, temp12;
        
        if (j == 0) {   
          if (isOneOpen) {
            temp9 = ca1DescCol;
            temp10 = ca1RespCol;
            temp11 = ca1DueCol;
            temp12 = ca1CompCol
          } else if (isTwoOpen) {
            temp9 = ca2DescCol;
            temp10 = ca2RespCol;
            temp11 = ca2DueCol;
            temp12 = ca2CompCol
          } else  if (isThreeOpen) {
            temp9 = ca3DescCol;
            temp10 = ca3RespCol;
            temp11 = ca3DueCol;
            temp12 = ca3CompCol
          } else {
            temp9 = ca4DescCol;
            temp10 = ca4RespCol;
            temp11 = ca4DueCol;
            temp12 = ca4CompCol
          }
        }
        
        if (j == 1) {
          if (isOneOpen && isTwoOpen) {
            temp9 = ca2DescCol;
            temp10 = ca2RespCol;
            temp11 = ca2DueCol;
            temp12 = ca2CompCol
          } else  if (isThreeOpen) {
            temp9 = ca3DescCol;
            temp10 = ca3RespCol;
            temp11 = ca3DueCol;
            temp12 = ca3CompCol
          } else {
            temp9 = ca4DescCol;
            temp10 = ca4RespCol;
            temp11 = ca4DueCol;
            temp12 = ca4CompCol
          }
        }
        
        if (j == 2) {
          if (isThreeOpen) {
            temp9 = ca3DescCol;
            temp10 = ca3RespCol;
            temp11 = ca3DueCol;
            temp12 = ca3CompCol
          } else {
            temp9 = ca4DescCol;
            temp10 = ca4RespCol;
            temp11 = ca4DueCol;
            temp12 = ca4CompCol
          }
        }
        
        if (j == 3) {
            temp9 = ca4DescCol;
            temp10 = ca4RespCol;
            temp11 = ca4DueCol;
            temp12 = ca4CompCol
          }
        
        output[0][9] = data[i][temp9]; //description
        output[0][10] = data[i][temp10]; //responsible party
        output[0][11] = data[i][temp11]; //due date
        output[0][12] = data[i][temp12]; //completion date
        
        var outputRow = outputSheet.getLastRow() + 1;
        outputSheet.getRange(outputRow,1,1,13).setValues(output);
        
        var now = new Date();
        if (output[0][11] < now) {
          pastDue(output);
        }
      }
    }
  }
}


// this function pastes the past due audit data into the target sheet
function pastDue(output) {
  var useRow = pastdueSheet.getLastRow() + 1;
  pastdueSheet.getRange(useRow,1,1,13).setValues(output);
  if (emailFlag) {
    getPastDueEmail(output);
  }
}

// find email addresses for responsible parties on past due audits
function getPastDueEmail(data) {
  var emailList = emailSheet.getRange("A3:B").getValues();
  for (var i=0;i<emailList.length;i++) {
    if (emailList[i][0] === data[0][10]) {
      sendPastDue(emailList[i][1],data);
    }
  }
}

// run a date check and return whether it is more than 30 days old
function checkDate(cell) {
  var now = new Date();
  var thirtyDays = new Date(now.getTime() - (30 * (1000 * 60 * 60 * 24)));
  if ((cell instanceof Date) && (cell < thirtyDays)) {
    return true;
  }
  return false;
}

// check the input row to see if there are any corrective actions not listed as complete
function checkOpen(exist,complete) {
  if (exist != '' && complete == '') {
    return true;
  }
  return false;
}

// find columns based on labels and returns the column number
function findColumn(criteria,data) {
  for (i=1;i<=data[0].length;i++) {
    if (data[0][i]) {
      var test = (data[0][i].trim());
    } else {var test = ""};
    if (data[0][i] == criteria || test == criteria) {
      return i;
    } 
  }
  return;
}

// sort a range in a sheet based on specified column
function sortSheet(sheet) {
  var range = sheet.getRange("A2:M"); //set this to the range you want included in the sort
  range.sort({column: 11, ascending: true});
}

// clears a sheet from row 2 to the end
function clearSheet(sheet) {
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var range = sheet.getRange(2,1,lastRow,lastColumn);
  range.clear();
}

// this function sends the regular summary email
function sendEmail() {
  var email = emailSheet.getRange("D2").getValue();
  var replyTo = emailSheet.getRange("D6").getValue();
  var url = 'https://docs.google.com/a/whirlpool.com/spreadsheets/d/1R-D7gBguUtWntgli6aq-CSYnWYZ_yXYopcHnnjzk4ds/edit?usp=sharing';
  var subject = 'Oustanding SSA Corrective Actions';
  var body = 'Hello,\n'+
    'The list of open SSA corrective actions older than 30 days is updated.\n\n'+
    url+'\n\n'+
      'Please take a few minutes to read through the update and work to close out open items.\n'+
      'Be aware that replying to this email will send an email to:'+replyTo+'\n'+  
      'Thank you.\n'+
      '(this is an automated message, if errors please notify your facilitator).';
  
//  Logger.log(email+' '+replyTo+' '+subject+' '+body);
  try {
    MailApp.sendEmail(email, replyTo, subject, body);
  } catch (e) {
    logError(e);
  }
}

// this function sends the past due email for each item
function sendPastDue(email,data) {
  var audit = data[0][1];
  var line = data[0][2];
  var area = data[0][3];
  var date = data[0][5];
  var observer = data[0][7];
  var description = data[0][9];
  var name = data[0][10];
  var dueDate = data[0][11];
  var replyTo = emailSheet.getRange("D9").getValue();
  var environmentAudit = emailSheet.getRange("E13").getValue();
  var equipmentAudit = emailSheet.getRange("E14").getValue();
  var materialsAudit = emailSheet.getRange("E15").getValue();
  var walkingAudit = emailSheet.getRange("E16").getValue();
  var url;
  
  if (audit == 'Environment') {
    url = environmentAudit;
  } else if (audit == 'Equipment') {
    url = equipmentAudit;
  } else if (audit == 'Materials') {
    url = materialsAudit;
  } else if (audit == 'Walking') {
    url = walkingAudit;
  }
  var subject = 'Past Due SSA Corrective Actions';
  var body = 'Hello, '+name+'\n'+
    'You have been listed as the corrective action(s) responsible party for a Standardized Safety Audit – '+audit+', line item '+line+'.\n'+
    'conducted by '+observer+' on '+date+' for the area of '+area+'.  The open corrective action is listed as: '+'\n'+description+'.\n'+
    'If the corrective actions have been completed please update the audit response spreadsheet here: \n'+url+'\nto reflect the date it was closed.'+
    'If you need the date extended or if you need help closing this open corrective action please contact '+replyTo+'\n'+
    'Be aware that replying to this email will send an email to:'+replyTo+'\n'+  
    'Thank you.\n'+
    '(this is an automated message, if errors please notify your facilitator).';
  
//  Logger.log(email+' '+replyTo+' '+subject+' '+body);
  try {
    MailApp.sendEmail(email, replyTo, subject, body);
  } catch (e) {
    logError(e);
  }
}


// the functions below are for email error handling
function logError(e) {
  var url = activeUrl();
  var email = "william_f_steinberger@whirlpool.com";
  
  MailApp.sendEmail(email, "Error report", 
                    "\r\nMessage: " + url
                    + "\r\nMessage: " + e.message
                    + "\r\nFile: " + e.fileName
                    + "\r\nLine: " + e.lineNumber);
  
  var errorSheet = SpreadsheetApp.openById('1AQxAvsMs6LF3qgFQFeVaSKIZr6xEV4YRMPYXHFhPRgI').getSheetByName('Errors');
  lastRow = errorSheet.getLastRow();
  var cell = errorSheet.getRange('A1');
  cell.offset(lastRow, 0).setValue(url);
  cell.offset(lastRow, 1).setValue(e.message);
  cell.offset(lastRow, 2).setValue(e.fileName);
  cell.offset(lastRow, 3).setValue(e.lineNumber);
}
function activeUrl() {
  var url, ss, doc, form;
  
  ss = SpreadsheetApp.getActiveSpreadsheet();
  doc = DocumentApp.getActiveDocument();
  form = FormApp.getActiveForm();
  
  if (ss != null && ss != undefined)
    url = ss.getUrl();
  else if (doc != null && doc != undefined)
    url = doc.getUrl();
  else if (form != null && form != undefined)
    url = form.getUrl();
  Logger.log(url);
  return url;
}

// this is a test function used to debug the past due emails without rebuilding the whole spreadsheet.  It is not normally used
function testPastDue() {
  var output = [[]];
  var outputData = pastdueSheet.getDataRange().getValues();
  for (var i=1;i<outputData.length;i++) {
    output[0] = outputData[i];
    getPastDueEmail(output);
  }
}


