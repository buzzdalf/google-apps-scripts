/*
* Possible revisions to make:
* - modify highlighting to only highlight if date requirement not met for that signature
* - send an email to area people when 7 day requirement not filled (data at least in 1st row of countermeasure/action)
* - send an email to area people when due date is reached
* - send an email to h&s admins when all sign-offs are in
* see deck for help: https://docs.google.com/presentation/d/1RCoc-GXvctmT-lgHFiPSVIJ9_UDyZVS6TxYMkvBxJ1g/edit#slide=id.g2023974555_0_7
* tied to: https://docs.google.com/spreadsheets/d/13_3VwQDfvVM53kDgTj8pxXd3jDwlCzaUAjURAxvQP2k
*/

function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('Update Reports')
  .addItem('Refresh Report Data', 'refresh')
  .addSeparator()
  .addToUi();
}

function refresh() {
  var sewoFolder = '0BxiuqJYDWNpkWVRmNFd0MEJYMHc';
  var folder = DriveApp.getFolderById(sewoFolder);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var files = getDriveFiles(folder);
  var fields = getFields(files);
  
  pasteFiles(ss, fields);
  populateTabs(ss, fields);
}

function getFields(files) {
  var fields = [];
  for (var i=0;i<files.length;i++) {
    var id = files[i].id;
    var url = files[i].url;
    var name = files[i].name;
    var tempSS = SpreadsheetApp.openById(id);
    var tempSheet = tempSS.getSheetByName("NAR SEWO Form");
    if (tempSheet) {
      var data = tempSheet.getDataRange().getValues();
      Logger.log(url);
      var fatality = data[2][4];
      var lostTime = data[3][4];
      var recordable = data[4][4];
      var firstAid = data[5][4];
      var nearMiss = data[6][4];
      var condition = data[7][4];
      var act = data[8][4];
      var what = data[11][1];
      var dept = data[3][8];
      var date = data[5][13];
      var doRow = findRow(data, "COUNTERMEASURE/ACTIONS", 1,url) + 1;
      var doBottom = findRow(data, "RESULTS ACHIEVED",1,url);
      var checkRow = findRow(data, "Check Performed by", 7,url) +1;
      var actRow = findRow(data, "Specific Expansion Areas", 1,url) +1;
      var reportingRow = findRow(data, "Name & Signature Injured Person/Person Reporting Issue", 1,url) + 1;
      var supervisorRow = findRow(data, "Name & Signature - Supervisor", 1,url) + 1;
      var plantLeadRow = findRow(data, "Name & Signature - Plant Leader",11,url) + 1;

      var inDo = checkSection(data,doRow,doBottom,1,6);
      var inCheck = checkSection(data,checkRow,checkRow+1,7,9);
      var inAct = data[actRow][1];
      var pdca = findPhase(inDo,inCheck,inAct);
      
      //var reporting = data[reportingRow][1];
      var reporting = checkSection(data,reportingRow,reportingRow+3,1,4) || '';
      //var ehs = data[reportingRow][11];
      var ehs = checkSection(data,reportingRow,reportingRow+3,11,14) || '';
      var supervisor = "";
      var manager = "";
      var director = "";
      var plantLead = "";

      if (data.length > supervisorRow) {
        var botSupRow = data.length;
        if (data.length > supervisorRow + 3) {
          botSupRow = supervisorRow + 3;
        }
        //supervisor = data[supervisorRow][1];
        supervisor = checkSection(data,supervisorRow,botSupRow,1,4) || '';
        if (plantLeadRow == supervisorRow) {
          //manager = data[reportingRow][5];
          manager = checkSection(data,reportingRow,reportingRow+3,5,10) || '';
          //director = data[supervisorRow][5];
          director = checkSection(data,supervisorRow,botSupRow,5,10) || '';
        } else {
          //manager = data[supervisorRow][5];
          manager = checkSection(data,supervisorRow,botSupRow,5,10) || '';
          //director = data[supervisorRow][11];
          director = checkSection(data,supervisorRow,botSupRow,11,14) || '';
        }
      }
      if (data.length > plantLeadRow) {
        var botLeadRow = data.length;
        if (data.length > plantLeadRow + 3) {
          botLeadRow = plantLeadRow + 3;
        }
        //plantLead = data[plantLeadRow][11];
        plantLead = checkSection(data,plantLeadRow,botLeadRow,11,14) || '';
      }          
    
      fields.push({url: url, name: name, fatality: fatality, lostTime: lostTime, recordable: recordable, firstaid: firstAid, 
                   nearmiss: nearMiss, condition: condition, act: act, what: what, dept: dept, date: date, 
                   inDo: inDo, inCheck: inCheck, inAct: inAct, pdca: pdca,
                   reporting: reporting, supervisor: supervisor, manager: manager, director: director, ehs: ehs, plantlead: plantLead});
    }
  }
//  Logger.log(fields);
  return fields;
}

function checkSection(data,doRow,doBottom,firstCol,lastCol) {
  for (var i=doRow;i<doBottom;i++) {
    for (var j=firstCol;j<lastCol;j++) {
      if (data[i][j]) {
        return data[i][j];
      }
    }
  }
}

function findPhase(inDo,inCheck,inAct) {
  var pdca = 'plan';
  if (inDo) {
    pdca = 'do';
    if (inCheck) {
      pdca = 'check';
      if (inAct) {
        pdca = 'act';
      }
    }
  }
  return pdca;
}

function pasteFiles(ss, files) {
  var sheet = ss.getSheetByName("all files");
  sheet.clear();
  var labels = [];
  labels.push(['url','name','date','dept','what','PDCA Step','fatality','lostTime','recordable','firstaid','nearmiss',
               'condition','act','reporting','supervisor','manager','director','ehs','plantlead']);
  
  for (var i=0;i<files.length;i++) {
    labels.push([files[i].url,files[i].name,files[i].date,files[i].dept,files[i].what,files[i].pdca,files[i].fatality,files[i].lostTime,
                 files[i].recordable,files[i].firstaid,files[i].nearmiss,files[i].condition,files[i].act,
                files[i].reporting,files[i].supervisor,files[i].manager,files[i].director,files[i].ehs,files[i].plantlead]);
  }
//  Logger.log(labels);
  sheet.getRange(1,1,labels.length,labels[0].length).setValues(labels);
  formatSheet(sheet);
  
}

function populateTabs(ss, files) {    //clean up duplicate code in here, be smarter about these, use variables, and arrays, only call things 1 time, etc.
  var actSheet = ss.getSheetByName("Act/Condition");
  var firstAidSheet = ss.getSheetByName("First Aid");
  var recordableSheet = ss.getSheetByName("Recordable");

  actSheet.clear();
  firstAidSheet.clear();
  recordableSheet.clear();

  var actLabels = [];
  var firstAidLabels = [];
  var recordableLabels = [];

  actLabels.push(['url','name','date','dept','what','PDCA Step','condition','act','reporting',
                  'supervisor']);
  firstAidLabels.push(['url','name','date','dept','what','PDCA Step','firstaid','nearmiss','reporting',
                       'supervisor','manager']);
  recordableLabels.push(['url','name','date','dept','what','PDCA Step','fatality','lostTime','recordable',
                         'reporting','supervisor','manager','director','ehs','plantlead']);
  
  for (var i=0;i<files.length;i++) {
    if (files[i].condition.length>0 || files[i].act.length>0) {
      actLabels.push([files[i].url,files[i].name,files[i].date,files[i].dept,files[i].what,files[i].pdca,files[i].condition,files[i].act,
                files[i].reporting,files[i].supervisor]);
    } else if (files[i].firstaid.length>0 || files[i].nearmiss.length>0) {
      firstAidLabels.push([files[i].url,files[i].name,files[i].date,files[i].dept,files[i].what,files[i].pdca,files[i].firstaid,files[i].nearmiss,
                files[i].reporting,files[i].supervisor,files[i].manager]);
    } else if (files[i].fatality.length>0 || files[i].lostTime.length>0 || files[i].recordable.length>0) {
      recordableLabels.push([files[i].url,files[i].name,files[i].date,files[i].dept,files[i].what,files[i].pdca,files[i].fatality,files[i].lostTime,files[i].recordable,
                             files[i].reporting,files[i].supervisor,files[i].manager,files[i].director,files[i].ehs,files[i].plantlead]);
    }
  }
//  Logger.log('acts:'+actLabels);
//  Logger.log('first aid:'+firstAidLabels);
//  Logger.log('recordables:'+recordableLabels);

  actSheet.getRange(1,1,actLabels.length,actLabels[0].length).setValues(actLabels);
  formatSheet(actSheet);
  firstAidSheet.getRange(1,1,firstAidLabels.length,firstAidLabels[0].length).setValues(firstAidLabels);
  formatSheet(firstAidSheet);
  recordableSheet.getRange(1,1,recordableLabels.length,recordableLabels[0].length).setValues(recordableLabels);
  formatSheet(recordableSheet);
}


function formatSheet(sheet){
  sortSheet(sheet);
  sheet.setFrozenRows(1);
  highlight(sheet);
}

function sortSheet(sheet) {
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var range = sheet.getRange(2,1,lastRow,lastColumn);
  range.sort(3);
  
}

function highlight(sheet) {
  var data = sheet.getDataRange().getValues();
  var column = findColumn(data,'PDCA Step');

  for (var j=0;j<data.length;j++) {
    var range = sheet.getRange(j+1,column+1);
    var phaseColor = '';
    if (data[j][column] == 'plan') {
      phaseColor = 'lime';
    } else if (data[j][column] == 'do') {
      phaseColor = 'blue';
    } else if (data[j][column] == 'check') {
      phaseColor = 'red';
    } else if (data[j][column] == 'act') {
      phaseColor = 'yellow';
    }
    range.setBackground(phaseColor);
  }
  
  var columnNames = ['reporting','supervisor','manager','director','ehs','plantlead'];
  for (var i=0;i<columnNames.length;i++) {
    var column = findColumn(data,columnNames[i]);
    if (column) {
      for (var j=0;j<data.length;j++) {
        if (!data[j][column]) {
          var range = sheet.getRange(j+1,column+1);
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

function findRow(data, criteria, column,sewo) {
  for (var i=1;i<=data.length;i++) {
//    Logger.log('sewo:'+sewo);
    try {
      if (data[i][column] == criteria) {
        return i;
      }
    } catch (e) {
//      logError(e);
      return;
    }
  }
  return;
}

// adapted from https://ctrlq.org/code/20034-search-drive-files
function getDriveFiles(folder) {
  // If Drive folder is not specified, start from the root folder
  if (folder == null) {
    return getDriveFiles(DriveApp.getRootFolder(), "");
  }
  
  var files = [];
  
  // Specify the MimeType of files you wish to search
  var fileIt = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while ( fileIt.hasNext() ) {
    var f = fileIt.next();
    files.push({id: f.getId(), url: f.getUrl(), name: f.getName()});
  }
  
  // Get all the sub-folders and iterate
  var folderIt = folder.getFolders();
  while(folderIt.hasNext()) {
    fs = getDriveFiles(folderIt.next());
    for (var i = 0; i < fs.length; i++) {
      files.push(fs[i]);
    }
  }
  
  return files;
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



//old code pasted in here that I might use

function tryMe(alert) {
  //var source = SpreadsheetApp.openById('1JMOtmGCfOdbxtUtmaXcW1yFN8BmgPZhhc1C4GGNsr_k');
  var source = SpreadsheetApp.openByUrl(alert.url);
  var sheet = source.getSheets()[0];
  var sheetName = sheet.getSheetName();
  
  sheet.copyTo(ss);
  
  var mailName = "Copy Of " + sheetName;
  var mailSheet = ss.getSheetByName(mailName);
  var mailID = mailSheet.getSheetId();
  
  var fatality = mailSheet.getRange("E3").getValue();
  var lostTime = mailSheet.getRange("E4").getValue();
  var recordable = mailSheet.getRange("E5").getValue();
  
  var noteRange = mailSheet.getRange("E3:E5");
  noteRange.clearNote();
  
  var pdfName = "Clyde - " + alert.name;
  
  if (fatality != '' || lostTime != '' || recordable != '') {
    
//    Logger.log('found one');
    
    clearRanges(mailSheet); 
    
    var fileName = savePDFs(destID, mailID, pdfName );
  }
  
  ss.deleteSheet(mailSheet);
  
  return fileName;
}


function sendEmail(myFiles) {
  var emailSheet = ss.getSheetByName('Standard E-Mail Text');
  var body = emailSheet.getRange("A4").getValue();
  var subject = emailSheet.getRange("A2").getValue();
  var emails = findEmails(emailSheet);
  
  for (var i=0;i<myFiles.length;i++) {
    body += ' '+myFiles[i]+'\n';
  }

//  Logger.log(emails+' '+subject+' '+ body);//+' '+myFiles);
  try {
    MailApp.sendEmail(emails, subject, body);//, {attachments: myFiles});
  } catch (e) {
    logError(e);
  }
}


