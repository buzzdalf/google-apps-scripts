/*
* Script to autmoatically email folks the Clyde Clips
* on a scheduled basis
* Setup Trigger to run runTomorow() Daily at 9 PM to look for tomorrow's date)
* tied to: https://docs.google.com/spreadsheets/d/127oePQkBXvBYdCgG_0US9bRspLdluW2Es5mzETP3BpQ
* Last edit 6/12/17 by: Bill Steinberger
*/

function onOpen() {
SpreadsheetApp.getUi()
.createMenu('Send E-mails')
.addItem('Run Email Script for Yesterday', 'runYesterday')
.addItem('Run Email Script for Today', 'runToday')
.addItem('Run Email Script for Tomorrow', 'runTomorrow')
.addToUi();
}

function runYesterday() {
  findEmails(-1);
}

function runToday() {
  findEmails(0);
}

function runTomorrow() {
  findEmails(1);
}

// this function extracts all the emails from a list in a spreadsheet and calls a function to send an email to each one
function findEmails(days) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet= ss.getSheetByName('emails');
  var data = sheet.getDataRange().getValues();
  var start = findList(data);
  var file = getClips(days);
  
//  Logger.log(file);
  
  if (file != null) {
    var emailArray = [];
    for (var i=start; i<data.length; i++) {
      var cell = [i,1];
      var email = data[i][0];
      if (email.indexOf('@') !== -1) {
        //sendEmail(email,file);
        emailArray.push(email);
      }
    }
    sendEmail(emailArray,file);
    sentClips(file);
  } else {
    var ssUrl = ss.getUrl();
    nothingFound(ssUrl);
  }
}

// this function finds the first row of a list in a spreadsheet, based on providing a name for the list heading
function findList(data) {
  var listHeading = 'Distribution list:';

  for (var i=0; i<data.length; i++) {
    if (data[i][0] == listHeading) {
      return (i+1);
    }
  }
  return;
}

//this function finds a file in a specific folder and checks to see if the file has today's day and month in the title
function getClips(days) {
  var folderID = '0B6zO-pLqjfSNY1MtZzFtZFVzVWc'; //https://drive.google.com/drive/folders/0B6zO-pLqjfSNY1MtZzFtZFVzVWc?usp=sharing
  var folder = DriveApp.getFolderById(folderID);
  var contents = folder.getFiles();
  
  while (contents.hasNext()) {
    var date = {};
    var file = contents.next();
    var fileName = file.getName();
    var fileId = file.getId();
    var fileUrl = file.getUrl();
    date = getDate(days);
//    Logger.log (fileName+' '+(date.date)+' '+fileName.indexOf(date.date)+' '+date.string+' '+fileName.indexOf(date.string));
    if ((fileName.indexOf(date.date) > 7 && fileName.indexOf(date.date) < 12) || (fileName.indexOf(date.string) > 7 && fileName.indexOf(date.string) < 12)){ 
//      Logger.log('true');
      return fileUrl;      
    }
//    Logger.log('false');
    return null;
  }
}

// this function gets today's date and returns the day and month
function getDate(days) {
  var fullDate = {};
  var now = new Date(); 
  var tomorrow = new Date(now.getTime() + (days * (1000 * 60 * 60 * 24))); //changed script to look for TOMORROW'S date so I can run the script on Sunday night vs Monday morning
  var day = tomorrow.getDate();
  var month = tomorrow.getMonth()+1;
  fullDate.date = month + "/" + day;
  if (day < 10) {
    fullDate.string = month + '/0'+day;
  }

  return fullDate;
}

// this function actually sends the email
function sendEmail(email,url) {
  var subject = 'Clyde Clips is out!';
  var body = "Hello, "+'\n'+
    "Here is a link to this week's Clyde Clips for your reading pleasure:"+'\n'+'\n'+
      url+'\n'+'\n'+
        "Please take a few minutes to read through this week's update and share it with your team."+'\n'+
          "Thank you."+'\n'+
        "(this is an automated message, if errors please simply notify your facilitator).";

  //Logger.log(email+' '+subject+' '+body);
  MailApp.sendEmail(email, subject, body);
}

function sentClips(url) {
  var email = 'william_f_steinberger@whirlpool.com';
  var subject = 'Clyde Clips Sent';
  var body = "Hello, "+'\n'+
    "The Clyde Clips Script ran today.  Here is what was sent:"+'\n'+'\n'+
      url+'\n'+'\n'+
          "Thank you."+'\n'+
        "(this is an automated message, if errors please simply notify your facilitator).";
  MailApp.sendEmail(email, subject, body);
}

function nothingFound(url) {
  var email = 'william_f_steinberger@whirlpool.com';
  var subject = 'No Clyde Clips Today';
  var body = "Hello, "+'\n'+
    "The Clyde Clips Script ran today, but there was no file found to send.  Here is the folder I looked at:"+'\n'+'\n'+
      'https://drive.google.com/drive/folders/0B6zO-pLqjfSNY1MtZzFtZFVzVWc?usp=sharing'+'\n'+
        'Here is the script spreasheet in case you need to run it manually:'+'\n'+
          url+'\n'+'\n'+
          "Thank you."+'\n'+
        "(this is an automated message, if errors please simply notify your facilitator).";
  MailApp.sendEmail(email, subject, body);
}