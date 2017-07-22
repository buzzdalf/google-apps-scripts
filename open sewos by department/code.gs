/*
* This script is tied to: https://docs.google.com/spreadsheets/d/1H7OxUW38sHSS4Xuu5LxBHpYjY6ZWsbpaH73s9BMbAPg/edit#gid=0
*/

function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('Update List')
  .addItem('Refresh Open SEWOs', 'getResults')
  .addSeparator()
  .addItem('Send the e-mail', 'sendEmail')
  .addToUi();
}


function getResults() {
  var ss = SpreadsheetApp.openById('13_3VwQDfvVM53kDgTj8pxXd3jDwlCzaUAjURAxvQP2k');

  var actSheet = ss.getSheetByName("Act/Condition");
  var firstAidSheet = ss.getSheetByName("First Aid");
  var recordableSheet = ss.getSheetByName("Recordable");
  
  var actData = actSheet.getDataRange().getValues();
  var firstAidData = firstAidSheet.getDataRange().getValues();
  var recordableData = recordableSheet.getDataRange().getValues();
  
  materialsOpen(actData,firstAidData,recordableData);
  
}

function materialsOpen(act,firstAid,recordable) {
  var ss = SpreadsheetApp.openById('1H7OxUW38sHSS4Xuu5LxBHpYjY6ZWsbpaH73s9BMbAPg');
  var sheet = ss.getSheetByName("Opens");

  sheet.clear();

  var labels = [];
  labels.push(['type','url','name','date','dept','what','reporting','supervisor','manager','director','ehs','plantlead']);
  sheet.getRange(1,1,labels.length,labels[0].length).setValues(labels);
  
  getActs(sheet,act);
  getFirstAids(sheet,firstAid);
  getRecordables(sheet,recordable);
  
}

function getActs(sheet,act) {
  var fields = [];
  for (var i=1;i<act.length;i++) {
    var testDept = act[i][3].toString();
    if(testDept.indexOf("4") == 0 && (!act[i][8] || !act[i][9])){
      fields.push(['unsafe act',act[i][0],act[i][1],act[i][2],act[i][3],act[i][4],act[i][8],act[i][9]]);
     }
  }
  var myRow = pasteData(sheet,fields);
  var lastRow = fields.length;

  var columnNames = ['reporting','supervisor'];
  highlight(sheet,columnNames,myRow,lastRow);
}

function getFirstAids(sheet,firstAid) {
  var fields = [];
  for (var i=1;i<firstAid.length;i++) {
    var testDept = firstAid[i][3].toString();
    if(testDept.indexOf("4") == 0 && (!firstAid[i][8] || !firstAid[i][9] || !firstAid[i][10])){
      fields.push(['firstaid',firstAid[i][0],firstAid[i][1],firstAid[i][2],firstAid[i][3],firstAid[i][4],firstAid[i][8],firstAid[i][9],firstAid[i][10]]);
    }
  }
  var myRow = pasteData(sheet,fields);
  var lastRow = fields.length;
  
  var columnNames = ['reporting','supervisor','manager'];
  highlight(sheet,columnNames,myRow,lastRow);
}

function getRecordables(sheet,recordable) {
  var fields = [];
  for (var i=1;i<recordable.length;i++) {
    var testDept = recordable[i][3].toString();
    if(testDept.indexOf("4") == 0 && (!recordable[i][9] || !recordable[i][10] || !recordable[i][11] || !recordable[i][12] || !recordable[i][13] || !recordable[i][14])) {
      fields.push(['recordable',recordable[i][0],recordable[i][1],recordable[i][2],recordable[i][3],recordable[i][4],recordable[i][9],recordable[i][10], 
                 recordable[i][11],recordable[i][12],recordable[i][13],recordable[i][14]]);    
    }
  }
  var myRow = pasteData(sheet,fields);
  var lastRow = fields.length;
  
  var columnNames = ['reporting','supervisor','manager','director','ehs','plantlead'];
  highlight(sheet,columnNames,myRow,lastRow);
}

function pasteData(sheet,fields) {
  var myRow = sheet.getLastRow() + 1;
  sheet.getRange(myRow,1,fields.length,fields[0].length).setValues(fields);
  return myRow;
}

function highlight(sheet,columnNames,startRow,endRow) {
  var firstRow = sheet.getRange(1,1,1,12).getValues();
  var data = sheet.getRange(startRow,1,endRow,12).getValues();

  for (var i=0;i<columnNames.length;i++) {
    var column = findColumn(firstRow,columnNames[i]);
    if (column) {
      for (var j=0;j<data.length;j++) {
        if (!data[j][column]) {
          var range = sheet.getRange(j+startRow,column+1);
          range.setBackground('pink');
        }
      }
    }
  }
}

function findColumn(data, criteria) {
  for (var i=1;i<=data[0].length;i++) {
    if (data[0][i] == criteria) {
      return i;
    }
  }
  return;
}

function sendEmail() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var emailSheet = ss.getSheetByName('E-Mail');
  var emails = emailSheet.getRange("A2").getValue();
  var subject = emailSheet.getRange("A5").getValue();
  var body = emailSheet.getRange("A8").getValue();

  var myFile = ss.getUrl();
  body = body + '\n\n' + myFile;
  
//  Logger.log(emails+' '+subject+' '+ body);
  try {
    MailApp.sendEmail(emails, subject, body);
  } catch (e) {
    logError(e);
  }
}

// the functions below are for error handling
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



