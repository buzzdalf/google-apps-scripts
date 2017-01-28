/*
* Script to autmoatically populate the weekly celebrate information
* on both the summary sheet and google doc
* Setup triggers to run runFirst() at 4:00 AM and runSecond() at 5:00 AM
* This script is tied to: https://docs.google.com/spreadsheets/d/1YVpdQnlTKAsNE-OukTLW9GoOOa75cYiXlUxiLuLRKrQ
* by: Bill Steinberger
*/

//this section of code works in spreadsheet
var ss = SpreadsheetApp.getActiveSpreadsheet();
var results = [];
var doc = 'weekly celebrate.pdf';

function runFirst() {
  findNames();
  moveList();
  prepDoc();
  var url = savePdf(doc);
}

function runSecond() {
  findNames();
  moveList();
  prepDoc();
  var url = savePdf(doc);
  sendEmail(url);
}

function findNames() {
  var sheet= ss.getSheetByName('Form Responses 1');
  var data = sheet.getDataRange().getValues();
  var now = new Date();
  var lastWeek = new Date(now.getTime() - (7 * (1000 * 60 * 60 * 24)));
  var j = 0;
  
  for (var i=0;i<data.length;i++) {
    if (data[i][0] >= lastWeek) {
      j++;
      results[j] = data[i];
    }
  }
}

function moveList() {
  var sheet= ss.getSheetByName('Summary');
  var lastRow = sheet.getLastRow();

  sheet.getRange(4,1,lastRow,2).clear();
  for (var i=1;i<results.length;i++) {
    sheet.getRange(i+3,1).setValue(results[i][2]);
    sheet.getRange(i+3,2).setValue(results[i][3]);
  }
  sheet.getRange(4,1,lastRow,2).setWrap(true);  
}

//code below work in document
// var folderID = '0BxiuqJYDWNpkV3FBSmVJZHZuWXM';
var folder = DriveApp.getFolderById('0BxiuqJYDWNpkV3FBSmVJZHZuWXM');
var fileID = '1In6MYeQC4CPqtWiWgCKT7HO6BLa8dbNtp2FrdeWAVX4';

function prepDoc() {
  var doc = DocumentApp.openById(fileID);
  var docBody = doc.getBody();
  var nameStyle = {};
  nameStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Lobster';
  nameStyle[DocumentApp.Attribute.FONT_SIZE] = 48;
  var actionStyle = {};
  actionStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Yanone Kaffeesatz';
  actionStyle[DocumentApp.Attribute.FONT_SIZE] = 36;
   
  docBody.clear();
  for (var i=1;i<results.length;i++) {
    var name = docBody.appendParagraph(results[i][2]);
    var action = docBody.appendParagraph(results[i][3]);
    name.setAttributes(nameStyle);
    action.setAttributes(actionStyle);
    docBody.appendPageBreak();
  }
  return;
}

function savePdf(doc) {
  var url = 'https://docs.google.com/document/d/'+fileID;
  var url_ext = '/export?exportFormat=pdf&format=pdf';   //export as pdf
//  + '&gid=' + sheet.getSheetId()   //the sheet's Id
//  // following parameters are optional...
//  + '&size=letter'      // paper size
//  + '&portrait=true'    // orientation, false for landscape
//  + '&fitw=true'        // fit to width, false for actual size
//  + '&sheetnames=false&printtitle=false&pagenumbers=false'  //hide optional headers and footers
//  + '&gridlines=false'  // hide gridlines
//  + '&fzr=false';       // do not repeat row headers (frozen rows) on each page
  
  var options = {
    headers: {'Authorization': 'Bearer ' +  ScriptApp.getOAuthToken()}
  }
 
  deleteFile(doc);

  var response = UrlFetchApp.fetch(url + url_ext, options);
  var blob = response.getBlob().setName(doc);
  var file = folder.createFile(blob);
  var fileUrl = file.getUrl();

  return fileUrl;
}

function deleteFile(fileName) {
  //deleted existing files so new ones can be created, maintaining only the latest copy
  var files = folder.getFilesByName(fileName);
  while (files.hasNext()) {
    var file = files.next();
    var fileId = file.getId();
    Drive.Files.remove(fileId);
  }
}

function sendEmail(url) {
  //send email out to folks so they know the file is ready to use
  var email = 'heather_r_kastor@whirlpool.com,william_f_steinberger@whirlpool.com';
  var subject = 'Weekly Safety Celebrate Updated';
  var body = "Hello, "+email+'\n'+'\n'+
    "The weekly safety celebrate file is ready to put on Clyde TV"+'\n'+
      "Here is a link to the updated file for importing:"+'\n'+
        url+'\n'+'\n'+
          "Here is a link to the folder the file is stored in in case you need that:"+'\n'+
          "https://drive.google.com/drive/u/0/folders/0BxiuqJYDWNpkV3FBSmVJZHZuWXM";
  MailApp.sendEmail(email, subject, body);
}

function forAuth_() {
  DriveApp.getFileById("Just for authorization"); // https://code.google.com/p/google-apps-script-issues/issues/detail?id=3579#c36
}
