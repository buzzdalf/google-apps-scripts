/*
* To automate emailing the results of this scripting, set a trigger to run pdfEmail
* This script is tied to: https://docs.google.com/spreadsheets/d/1Z9jJ0pVl7XvlkGe-ohWR55j-lE0jhMcdexFhzAPYDgg 
*/

var destID = '1Z9jJ0pVl7XvlkGe-ohWR55j-lE0jhMcdexFhzAPYDgg';
var ss = SpreadsheetApp.openById(destID);


function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('Export SEWOs')
  .addItem('Build pdfs Only', 'pdfOnly')
  .addSeparator()
  .addItem('Build pdfs and Send Email', 'pdfEmail')
  .addToUi();
}

function pdfOnly() {
  getSewos(false);
}

function pdfEmail() {
  getSewos(true);
}

function getSewos(sending) {
  var weekday = isWeekday();
  if (weekday) {
    var sewoFolder = '0BxiuqJYDWNpkWVRmNFd0MEJYMHc';
    var folder = DriveApp.getFolderById(sewoFolder);
    var search = 'title contains "SEWO" and mimeType contains "spreadsheet" and not title contains "Form Master" and not title contains "Export"';
    //  var search = 'mimeType contains "spreadsheet" and (modifiedDate > "' + startTime.start + '") and not title contains "Form Master" and not title contains "test"';
    var alerts = searchFolder(folder,search);
    var text = '';
    var files = [];
    if (alerts.length > 0) {
      for (var i=0;i<alerts.length;i++) {
        var file = tryMe(alerts[i]);
        if (file) {
          //        files.push(file.getAs(MimeType.PDF));
          files.push(file);
        }
        //      tryMe(alerts[i]);
      }
//      Logger.log(files);
      if (sending && (files.length > 0)) {
        sendEmail(files);
//        Logger.log('sending email');
      }
    }
  }
}

function tryMe(alert) {
  //var source = SpreadsheetApp.openById('1JMOtmGCfOdbxtUtmaXcW1yFN8BmgPZhhc1C4GGNsr_k');
  var source = SpreadsheetApp.openByUrl(alert.url);
  var sheet = source.getSheets()[0];
  var sheetName = sheet.getSheetName();
  
  sheet.copyTo(ss);
  
  var mailName = "Copy Of " + sheetName;
  var mailSheet = ss.getSheetByName(mailName);
  var mailID = mailSheet.getSheetId();
  
  var sewoTest = mailSheet.getRange("D3").getValue();
  var fatality = mailSheet.getRange("E3").getValue();
  var lostTime = mailSheet.getRange("E4").getValue();
  var recordable = mailSheet.getRange("E5").getValue();
  
  var noteRange = mailSheet.getRange("E3:E5");
  noteRange.clearNote();
  
  var pdfName = "Clyde - " + alert.name;
  
  if (sewoTest == 'Fatality' && (fatality != '' || lostTime != '' || recordable != '')) {
    
//    Logger.log('found one');
    
    clearRanges(mailSheet); 
    
    var fileName = savePDFs(destID, mailID, pdfName );
  }
  
  ss.deleteSheet(mailSheet);
  
  return fileName;
}

function clearRanges(mailSheet) {
  var range = ['G6','I6','B27:O34','B51:O60','B62:O66','B68:O70','B81:O90']
  
  for (var i=0;i<range.length;i++) {
    mailSheet.getRange(range[i]).clearContent();
  }
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

function isWeekday() {
  var today =  new Date();
  //Skip week-end 6=Sat, 0=Sun, 1=Mon
  if (today.getDay()==1 || today.getDay()==0) {
    return false;
  }
  return true;
}

// this function extracts all the emails from a list in a spreadsheet and calls a function to send an email to each one
function findEmails(sheet) {
  var sheet= ss.getSheetByName('E-Mail Disitribution List');
//  var sheet= ss.getSheetByName('Test E-Mail List'); //this line is for testing changes
  var data = sheet.getDataRange().getValues();
  var emailList = [];
  
  for (var i=0; i<data.length; i++) {
    var email = data[i][0];
    if (email.indexOf('@') !== -1) {
      emailList.push(email);
    }
  }
  return emailList;
}

/**
 * Export one or all sheets in a spreadsheet as PDF files on user's Google Drive,
 * in same folder that contained original spreadsheet.
 *
 * Adapted from https://code.google.com/p/google-apps-script-issues/issues/detail?id=3579#c25
 *
 * @param {String}  optSSId       (optional) ID of spreadsheet to export.
 *                                If not provided, script assumes it is
 *                                sheet-bound and opens the active spreadsheet.
 * @param {String}  optSheetId    (optional) ID of single sheet to export.
 *                                If not provided, all sheets will export.
 */
function savePDFs( optSSId, optSheetId, pdfName ) {

  // If a sheet ID was provided, open that sheet, otherwise assume script is
  // sheet-bound, and open the active spreadsheet.
  var ss = (optSSId) ? SpreadsheetApp.openById(optSSId) : SpreadsheetApp.getActiveSpreadsheet();

  // Get URL of spreadsheet, and remove the trailing 'edit'
  var url = ss.getUrl().replace(/edit$/,'');

  // Get folder containing spreadsheet, for later export
  var parents = DriveApp.getFileById(ss.getId()).getParents();
  if (parents.hasNext()) {
    var folder = parents.next();
  }
  else {
    folder = DriveApp.getRootFolder();
  }

  // Get array of all sheets in spreadsheet
  var sheets = ss.getSheets();

  // Loop through all sheets, generating PDF files.
  for (var i=0; i<sheets.length; i++) {
    var sheet = sheets[i];

    // If provided a optSheetId, only save it.
    if (optSheetId && optSheetId !== sheet.getSheetId()) continue; 

    //additional parameters for exporting the sheet as a pdf
    var url_ext = 'export?exportFormat=pdf&format=pdf'   //export as pdf
        + '&gid=' + sheet.getSheetId()   //the sheet's Id
        // following parameters are optional...
        + '&size=letter'      // paper size
        + '&portrait=true'    // orientation, false for landscape
        + '&fitw=true'        // fit to width, false for actual size
        + '&sheetnames=false&printtitle=false&pagenumbers=false'  //hide optional headers and footers
        + '&gridlines=false'  // hide gridlines
        + '&fzr=false';       // do not repeat row headers (frozen rows) on each page

    var options = {
      headers: {
        'Authorization': 'Bearer ' +  ScriptApp.getOAuthToken()
      }
    }

    var response = UrlFetchApp.fetch(url + url_ext, options);

    //    var blob = response.getBlob().setName(ss.getName() + ' - ' + sheet.getName() + '.pdf');
    
    var blob = response.getBlob().setName(pdfName + '.pdf');
    
    //from here you should be able to use and manipulate the blob to send and email or create a file per usual.
    //In this example, I save the pdf to drive
    //    folder.createFile(blob);
    var pdfFile = folder.createFile(blob);
    pdfFile.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);

    var url = pdfFile.getUrl(); 
    
  }
  //  return blob;
  return url;
}

/**
 * Dummy function for API authorization only.
 * From: http://stackoverflow.com/a/37172203/1677912
 */
function forAuth_() {
  DriveApp.getFileById("Just for authorization"); // https://code.google.com/p/google-apps-script-issues/issues/detail?id=3579#c36
}

function searchFolder(folder,search) {
  var files  = folder.searchFiles(search);
  var alerts = [];
  while (files.hasNext()) {
    var fileInfo = {};
    var file = files.next();
    var fileName = file.getName();
    var fileUrl = file.getUrl();
//    var type = file.getMimeType();
//    Logger.log(fileName+' '+type+' '+fileUrl);
    fileInfo = {
      name:fileName,
      url:fileUrl
    };
    alerts.push(fileInfo);
  }
  return alerts;
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
