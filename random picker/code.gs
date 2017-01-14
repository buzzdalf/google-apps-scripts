/* 
*  Script to automatically pick a random location for gemba walks
*  and notify the leadership in the area selected
*  also increments that area in the list to prevent repeats
*  Setup Triggers to run findGembas() every night at 1:00 AM
*  Tied to the following sheet: https://docs.google.com/spreadsheets/d/1q7J9aHqc70fXQtgmITpepv12ByyzRa6-YkOoS3j-Iyo
*  by: Bill Steinberger, please contact me with any issues or questions
*/

var usedColumn = 2;
var emailColumn = 3;
var countColumn = 4;
var totalColumns = 5;
var startRow = 10;
var resultRow = 7;
var sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("gemba");
var lastRow = sheet1.getLastRow();
var Avals = sheet1.getRange(startRow,1,lastRow,1).getValues();
var totalRows = Avals.filter(String).length;
//var data = sheet1.getRange(startRow,1,lastRow,totalColumns).getValues();
var count = 0;
var allTitles = [];

function picker() {
  var randNumber = Math.floor((Math.random() * totalRows) + 1);
  var randRow = randNumber+startRow-1;
  var winner = sheet1.getRange(resultRow,1);
  var rangeArray = sheet1.getRange(startRow, countColumn, totalRows, 1).getValues();
  var maxInRange = rangeArray.sort(function(a,b){return b-a})[0][0];
  var minInRange = rangeArray.sort(function(a,b){return a-b})[0][0];
  do {
    var randRow = randPick();
    var countChosen = sheet1.getRange(randRow,countColumn).getValue();
  } while (countChosen != minInRange);
  var chosenRow = sheet1.getRange(randRow,1,1,usedColumn);
  chosenRow.copyTo(winner);
  return randRow;
}

function randPick() {
  var randNumber = Math.floor((Math.random() * totalRows) + 1);
  var randRow = randNumber+startRow-1;
  return randRow;
}

function findGembas() {
  var now = new Date();
  var endDate = new Date(now.getTime() + (7 * (1000 * 60 * 60 * 24)));
  for (var i=1;i<8;i++) {
    var searchDate = new Date(now.getTime() + (i * (1000 * 60 * 60 * 24)));
    findEvent(searchDate,count);
  }
  Logger.log(count);
  if (count == 0) {
    showAlert();
  } else {
    sendSummary();
  }
}

function findEvent(date) {
  var eventTitle = 'GEMBA Walk (leadership)';
  var cal = CalendarApp.getDefaultCalendar();
  var events = cal.getEventsForDay(date);
//  var eventExist = false;
  for (var i=0;i<events.length;i++) {
    var details=[[events[i].getTitle()]];
    if (details == eventTitle) {
      //Logger.log('event'+details);
//      eventExist = true;
      count++;
      var pickedRow = updateTitle(events[i]);
      sendEmail(date,pickedRow);
      increment(date,pickedRow);
    }
  }
//  if (!eventExist) {
//    showAlert();
//  }
  return count;
}

function updateTitle(event) {
  var oldTitle = event.getTitle();
  oldTitle = oldTitle.slice(0, -13);
  var pickedRow = picker();
  var itemTitle = sheet1.getRange(pickedRow,1).getValue();//data[cell[0]][0];
  allTitles.push(itemTitle);
  var newTitle = oldTitle + ' - ' + itemTitle;
  //Logger.log(newTitle);
  event.setTitle(newTitle);
  invite(event,pickedRow);
  return pickedRow;
}

function increment(fullDate,row) {
  var date = (fullDate.getMonth()+1) + "/" + fullDate.getDate() + "/" +  fullDate.getFullYear();
  var count = sheet1.getRange(row,4).getValue();
  count++;
  sheet1.getRange(row,4).setValue(count);
  //Logger.log("newCount:"+count);
  sheet1.getRange(row,5).setValue(date);
  //Logger.log("date:"+date);
}

function invite(event,row) {
  var emails = sheet1.getRange(row,3).getValue();//data[cell[0]][1];
  var email = emails.split(',');
  for (var i=0;i<email.length;i++) {
    //Logger.log(email[i]);
    event.addGuest(email[i]);
  }
}

function sendEmail(fullDate,row) {
  var date = (fullDate.getMonth()+1) + "/" + fullDate.getDate() + "/" +  fullDate.getFullYear();
  var email = sheet1.getRange(row,3).getValue();//data[cell[0]][1];
  var event = sheet1.getRange(row,1).getValue();;
  var subject = 'Notice of upcoming Gemba Walk';
  var body = "Hello, "+email+'\n'+
    "This is notification that you have a Gemba Walk coming up on "+date+'\n'+
    "Please plan to facilitate a Gemba walk in the "+event+" area at 7:45 AM on "+date+"."+'\n'+
    "If this date or time does not work for you, please notify me immediately so we can make other arrangements."+'\n'+
    "Thank you."+'\n'+
    "(this is an automated message, if errors please simply notify your Gemba facilitator).";
  //Logger.log(email+' '+subject+' '+body); 
  MailApp.sendEmail(email, subject, body);  
}

function sendSummary() {
  var email = 'william_f_steinberger@whirlpool.com';
  var subject = 'Gembas sent today:'+count;
  var body = "Hello, "+email+'\n'+
    "There were a total of "+count+" Gemba notifications sent today"+'\n'+
    "For the following topics: "+allTitles+'\n'+
    "(this is an automated message, if errors please simply notify your Gemba facilitator).";
  //Logger.log(email+' '+subject+' '+body); 
  MailApp.sendEmail(email, subject, body);  
}

function showAlert() {
  var email = 'william_f_steinberger@whirlpool.com';
  var subject = 'No unplanned Gembas';
  var body = "Hello, "+email+'\n'+
    "The Gemba notification system attempted to update the calendar"+'\n'+
    "But there are no un-planned Gembas in the specified date range."+'\n'+
    "(this is an automated message, if errors please simply notify your A-3 review facilitator).";
  MailApp.sendEmail(email, subject, body);  
}

