/*
* Tied to: https://docs.google.com/spreadsheets/d/1X-OTk-1GbwYIABKtAIOt21sBW_5T7nrRmhCF3gu8fVo
*/

function runAll() {
  var startTime = getStartTime();
  var incidents = getIncidents(startTime);
  var alerts = getAlerts(startTime);
  var sewo = getSewos(startTime);
  var data = dailyUpdate();
  var emails = findEmails();
  sendEmail(incidents, alerts, sewo, data, emails);
}

function getIncidents(startTime) {
  var incidentFolder = '0BxiuqJYDWNpkN0ZjUUpVWmJsLWc';
  var folder = DriveApp.getFolderById(incidentFolder);
  var search = 'mimeType contains "pdf" and (modifiedDate > "' + startTime.start + '")';
  var alerts = searchFolder(folder,search);
  var text = '';
  if (alerts.length > 0) {
    text = 'Clyde had '+alerts.length+' safety incidents in the last 24 hours.  The incidents are listed and attached here:'+'\n';
    for (var i=0;i<alerts.length;i++) {
      text += alerts[i].name+': '+alerts[i].url+'\n';
    }
  } else {
    text = 'Clyde did not have any new safety incidents in the last 24 hours.  Here is the folder I looked in:'+'\n'+
      'https://drive.google.com/drive/u/0/folders/'+incidentFolder+'\n';
  }
  return text;
}

function getAlerts(startTime) {
  var alertFolder = '0BxiuqJYDWNpkaVdxWU1pYjR3Mnc';
  var folder = DriveApp.getFolderById(alertFolder);
  //var search = 'title contains "Safety Alert" and title contains "'+startTime.year+'" and mimeType contains "spreadsheet" and (modifiedDate > "' + startTime.start + '")';
  var search = 'mimeType contains "spreadsheet" and (modifiedDate > "' + startTime.start + '")';
//  var search = 'title contains "Safety Alert"';

  var alerts = searchFolder(folder,search);
  var text = '';
  if (alerts.length > 0) {
    text = 'Clyde had '+alerts.length+' safety alerts in the last 24 hours.  The alerts are listed and attached here:'+'\n';
    for (var i=0;i<alerts.length;i++) {
      text += alerts[i].name+': '+alerts[i].url+'\n';
    }
  } else {
    text = 'Clyde did not have any new safety alerts in the last 24 hours. Here is the folder I looked in:'+'\n'+
      'https://drive.google.com/drive/u/0/folders/'+alertFolder+'\n';
  }
  return text;
}

function getSewos(startTime) {
  var sewoFolder = '0BxiuqJYDWNpkWVRmNFd0MEJYMHc';
  var folder = DriveApp.getFolderById(sewoFolder);
  //var search = 'title contains "SEWO" and title contains "'+startTime.month+'" and title contains "'+startTime.year+'" and mimeType contains "spreadsheet" and (modifiedDate > "' + startTime.start + '")';
  var search = 'mimeType contains "spreadsheet" and (modifiedDate > "' + startTime.start + '") and not title contains "Form Master" and not title contains "Export"';
  var alerts = searchFolder(folder,search);
//  Logger.log(startTime+' '+alerts);
  var text = '';
  if (alerts.length > 0) {
    text = 'Clyde had '+alerts.length+' SEWOs in the last 24 hours.  The SEWOs are listed and attached here:'+'\n';
    for (var i=0;i<alerts.length;i++) {
      text += alerts[i].name+': '+alerts[i].url+'\n';
    }
  } else {
    text = 'Clyde did not have any SEWOs in the last 24 hours. Here is the folder I looked in:'+'\n'+
      'https://drive.google.com/drive/u/0/folders/'+sewoFolder+'\n';
  }
  return text;
}

function getStartTime() {
  var today = new Date();
//  var weekday = isWeekday(today);
  var i = 1;
//  if (!weekday) {
//    i = 3;
//  }
  var yesterday = new Date(today.getTime() - i * 25 * 60 * 60 * 1000);
  var day = yesterday.getDate();
  var month = yesterday.getMonth()+1;
  var year = yesterday.getFullYear();
  var start = yesterday.toISOString();
  return {
    day:day,
    month:month,
    year:year,
    start:start
  };
}

function searchFolder(folder,search) {
  var files  = folder.searchFiles(search);
  var alerts = [];
  while (files.hasNext()) {
    var fileInfo = {};
    var file = files.next();
    var fileName = file.getName();
    var fileUrl = file.getUrl();
    var type = file.getMimeType();
//    Logger.log(fileName+' '+type+' '+fileUrl);
    fileInfo = {
      name:fileName,
      url:fileUrl
    };
    alerts.push(fileInfo);
  }
  return alerts;
}

function dailyUpdate() {
  var sheet = SpreadsheetApp.openById('1hotE5vQlYhnydMgUMdFLcVHMgza0kQQkjEpHrJGr1FU').getSheetByName('data');
  var data = sheet.getRange("A3:B8").getValues();
  var date = getDate(sheet.getRange("B1").getValue());
  var text = sheet.getRange("A1").getValue()+': '+date+'\n';
  for (var i=1;i<data.length;i++) {
    text += data[i][0]+': '+data[i][1]+'\n';
  }
  //Logger.log(text);
  return text;
}

// this function extracts all the emails from a list in a spreadsheet and calls a function to send an email to each one
function findEmails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet= ss.getSheetByName('emails');
  var data = sheet.getDataRange().getValues();
  var emailList = [];
  
  for (var i=0; i<data.length; i++) {
    var cell = [i,1]; //is this old, leftover code?  try removing on next edit
    var email = data[i][0];
    if (email.indexOf('@') !== -1) {
      emailList.push(email);
    }
  }
  return emailList;
}

// this function take's a date and returns the month / day / year
function getDate(now) {
  if (now instanceof Date) {
    var day = now.getDate();
    var month = now.getMonth()+1;
    var year = now.getFullYear();
    var fullDate = month + "/" + day + "/" + year; 
    return fullDate;
  }
}

function isWeekday(day) {
  //Skip week-end 6=Sat, 0=Sun
  if (day.getDay()==6 || day.getDay()==0) {
    return false;
  }
  return true;
}

function sendEmail(incidents, alerts, sewo, daily, email) {
  var dailySheet = 'https://docs.google.com/spreadsheets/d/1hotE5vQlYhnydMgUMdFLcVHMgza0kQQkjEpHrJGr1FU';
  var subject = 'Daily Clyde Safety Update';
  var body = 'Here is todays daily safety update for Clyde:'+'\n'+'\n'+
    daily+'\n'+
      'You can see the full report here:'+'\n'+
        dailySheet+'\n'+'\n'+
          'Note: The numbers above may be different that what is pulled from each folder based on timing for people entering reports, etc.'+'\n'+
            'Here is the safety alert update'+'\n'+
            alerts+'\n'+
              'Here is the list of recent incidents'+'\n'+
                incidents+'\n'+
              'Here is the list of recent SEWOs'+'\n'+
                sewo+'\n'+
                  'Thank you.'+'\n'+
                    '(this is an automated message, if errors please notify your facilitator).';
  
//  Logger.log(email+' '+subject+' '+ body);
  try {
    MailApp.sendEmail(email, subject, body);
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
//  Logger.log(url);
  return url;
}
