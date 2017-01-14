/* 
* Script to autmoatically create calendar invites/ descriptions for each a-3 item
* based on the emails and descriptions loaded in the spreadsheet
* Setup Trigger to run a3start() every day at 3:00 AM
* This sscript is tied to this spreadsheet: https://docs.google.com/spreadsheets/d/14wz5tz1WFzv65O4Fdpjla2m-2Fy-XTJ_vpSqJI2mB8U
* by: Bill Steinberger
*/

var sheet = SpreadsheetApp.getActiveSheet();
var data = sheet.getDataRange().getValues();
var cal = CalendarApp.getDefaultCalendar();

function a3start() {
  var count = 0;
  var now = new Date();
  var twoWeeks = new Date(now.getTime() + (10 * (1000 * 60 * 60 * 24)));
  //Logger.log(now+' dates '+twoWeeks);
  for (var i=1; i<data.length; i++) {
    for (var j=3; j<data[i].length;j+=2) {
      if ((data[i][j] instanceof Date) && (now < data[i][j]) && (data[i][j] < twoWeeks) && (data[i][j+1] != 'yes')) {
        //Logger.log(data[i][j]);
        var a3Date = data[i][j];
        var cell = [i,j];
        findEvent(a3Date,cell);
        count++;
      }
    }
  }
  sendSummary(count);
}

function findEvent(date,cell) {
  var eventTitle = 'A3 Review (Leadership)';
  var events = cal.getEventsForDay(date);
  var eventExist = false;
  for (var i=0;i<events.length;i++) {
    var details=[[events[i].getTitle()]];
    if (details == eventTitle) {
      //Logger.log('event'+details);
      eventExist = true;
      updateDesc(events[i],cell);
      invite(events[i],cell);
      sendEmail(date,cell);
      setFlag(cell);
    }
  }
  if (!eventExist) {
    //no current A-3 reviews on this date
    showAlert(date,cell);
  }
}

function updateDesc(event,cell) {
  var oldDesc = event.getDescription();
  var itemDesc = data[cell[0]][0];
  var newDesc = oldDesc + '\n' + itemDesc;
  //Logger.log(newDesc);
  event.setDescription(newDesc);
}

function invite(event,cell) {
  var emails = data[cell[0]][1];
  var email = emails.split(',');
  for (var i=0;i<email.length;i++) {
    //Logger.log(email[i]);
    event.addGuest(email[i]);
  }
}

function sendEmail(fullDate,cell) {
//  var fullDate = data[cell[0]][cell[1]];
  var date = (fullDate.getMonth()+1) + "/" + fullDate.getDate() + "/" +  fullDate.getFullYear();
  var email = data[cell[0]][1];
  var subject = 'Notice of upcoming A-3 review';
  var body = "Hello, "+email+'\n'+
    "This is notification that you have an A-3 review coming up on "+date+" for the following project:"+'\n'+'\n'+
    data[cell[0]][0]+'\n'+'\n'+
    "Please upload your A-3 to the following google drive folder:"+'\n'+
    "https://drive.google.com/a/whirlpool.com/folderview?id=0B5SJSwfFrqbBbF8zSlpDc3MzZ0U&usp=sharing"+'\n'+
    "and come prepared to present the project."+'\n'+
    "Thank you.  See you at the review meeting on "+date+"."+'\n'+
    "(this is an automated message, if errors please simply notify your A-3 review facilitator).";
  //Logger.log(email+' '+subject+' '+body); 
  MailApp.sendEmail(email, subject, body);  
}

function setFlag(cell) {
  var row = cell[0] + 1;
  var col = cell[1] + 2;
  //Logger.log(row+' '+col);
  sheet.getRange(row,col).setValue('yes');
}

function sendSummary(count) {
  var email = 'william_f_steinberger@whirlpool.com';
  var subject = 'A-3s sent today:'+count;
  var body = "Hello, "+email+'\n'+
    "There were a total of "+count+" A-3 notifications sent today"+'\n'+
    "(this is an automated message, if errors please simply notify your A-3 review facilitator).";
  MailApp.sendEmail(email, subject, body);  
}

function showAlert(fullDate,cell) {
  var date = (fullDate.getMonth()+1) + "/" + fullDate.getDate() + "/" +  fullDate.getFullYear();
  var email = 'william_f_steinberger@whirlpool.com';
  var subject = 'A-3 date issue';
  var body = "Hello, "+email+'\n'+
    "The A-3 notification system attempted to update a calendar event on:"+date+" for the following project:"+'\n'+
    data[cell[0]][0]+'\n'+
    "But there are no A-3 review calendar entries on this date.  Please review and take action"+'\n'+
    "https://drive.google.com/a/whirlpool.com/folderview?id=0B5SJSwfFrqbBbF8zSlpDc3MzZ0U&usp=sharing"+'\n'+
    "(this is an automated message, if errors please simply notify your A-3 review facilitator).";
  MailApp.sendEmail(email, subject, body);  
}