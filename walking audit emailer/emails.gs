/*
* tied to: https://docs.google.com/spreadsheets/d/1ZhsdQQrgOzSWsB-lmRYhg4PhOjsn6QpSIAgqkxTvxNI
*/

var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('Data of Form Responses');
var rawDataSheet = ss.getSheetByName('Form Responses 2');
var emailSheet = ss.getSheetByName('Email Template');
var emailTemplate = emailSheet.getRange("A1").getValue();

function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('Send E-mails')
  .addItem('Send e-mails to catch-up', 'catchUp')
  .addToUi();
}

function catchUp() {
  var rawData = rawDataSheet.getDataRange().getValues();
  var testColumn = findColumn(normalizeHeader('Timestamp'),rawData); //load in the column name for timestamp
  var sentColumn = findColumn(normalizeHeader('email sent?'),rawData); //load in the column name for email sent flag
  for (var i=1;i<rawData.length;i++) {
    var test = rawData[i][testColumn];
    var flag = rawData[i][sentColumn];
    if (test == '') {
      break;
    }
    if ( flag !== 'Yes') {
      addRow(i+1);
      var sent = buildEmail(i+1,emailTemplate);
      if (sent) {
        labelYes(i+1,rawData);
      }
    }
  }
}

function collectInfo(e) {
  var rawData = rawDataSheet.getDataRange().getValues();
  var testColumn = findColumn(normalizeHeader('Timestamp'),rawData); //load in the column name for timestamp
  for (var i=1;i<rawData.length;i++) {
    var test = rawData[i][testColumn];
    if (test == '') {
      break;
    }
  }
  addRow(i);
  var sent = buildEmail(i,emailTemplate);
  if (sent) {
    labelYes(i,rawData);
  }
}

function labelYes(key,rawData) {
  var sentColumn = findColumn(normalizeHeader('email sent?'),rawData); //load in the column name for email sent flag
  rawDataSheet.getRange(key, sentColumn+1).setValue('Yes');
}

// This function copies the formula down to the current audit row in the sheet
function addRow(row) {
  var formulas = sheet.getRange(row-1,1,1,10).getFormulasR1C1();
  sheet.getRange(row,1,1,10).setFormulasR1C1(formulas);
}

function buildEmail(key,template) {
  var data = sheet.getDataRange().getValues();
  var emailColumn = findColumn(normalizeHeader('Email Address'),data); //load in the column name for the emails to send to
  var emails = data[key-1][emailColumn];
  var emailText = fillInTemplate(template, data[key-1], data);
  
  var parties = [];
  var partyColumn = ['corrective action 1 responsible party',
                     'Corrective action 2 responsible party',
                     'Corrective Action 3 responsible party',
                     'Corrective Action 4 responsible party']; //load in the column names for the responsible party columns
  
  for (var k=0;k<partyColumn.length;k++) {
    var useColumn = findColumn(normalizeHeader(partyColumn[k]),data);
    var partyExist = data[key-1][useColumn];
    if (partyExist != '') {
      parties.push(partyExist);
    }
  }
   var sent = sendEmail(key,emails,emailText,parties);
  return sent;
}

function getEmail(list,name) {
  for (var i=0;i<list.length;i++) {
    if (list[i][0] === name) {
      return list[i][1];
    }
  }
}

function sendEmail(row,email,body,parties) {
  var setupSheet = ss.getSheetByName('Setup');
  var replyTo = setupSheet.getRange("B2").getValue();
  var subject = setupSheet.getRange("B3").getValue();
  var emailList = setupSheet.getRange("J28:K200").getValues();
  
  email += ',' + replyTo; //comment this out to stop getting all emails
  subject += ', row:' + row;
  
  for (var j=0;j<parties.length;j++) {
    var tempEmail = getEmail(emailList,parties[j]);
    email += ',' + tempEmail;
  }
  var plainBody = body.replace(/(<([^>]+)>)/ig, ""); // clear html tags for plain email
//  Logger.log(email+' '+replyTo+' '+subject+' '+ plainBody, {htmlBody:body});
  try {
    MailApp.sendEmail(email, subject, plainBody, {htmlBody:body, replyTo:replyTo});
    return true;
  } catch (e) {
    logError(e);
    return false;
  }
}

// find columns based on labels and returns the column number
function findColumn(criteria,dataRange) {
  for (i=1;i<=dataRange[0].length;i++) {
    if (dataRange[0][i]) {
      var test = normalizeHeader(dataRange[0][i]);
    } else { var test = "" };
    if (dataRange[0][i] == criteria || test == criteria) {
      return i;
    }
  }
  return;
}

function fillInTemplate(template, currentData, allData) {
  var email = template;
  // Search for all the variables to be replaced, for instance ${"Column name"}
  var templateVars = template.match(/\$\{\"[^\"]+\"\}/g);

  for (var i = 0; i < templateVars.length; ++i) {
    var checkLabel = normalizeHeader(templateVars[i].substr(3).slice(0, -2),allData); //this line crashes the script if there is no column match for the label
    var useColumn = findColumn(checkLabel,allData);
    var makeBold = '<span style="color:red;font-weight:bold">' + currentData[useColumn] + '</span>';
    email = email.replace(templateVars[i], makeBold || "");
  }
  return email;
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