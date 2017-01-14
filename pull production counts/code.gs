/*
Script to update the production tracking data
Run copyData script on schedule
There is also a script to go to today's date in the tracking sheet
* tied to: https://docs.google.com/spreadsheets/d/1CnMId9KbIvF5J1eW6Nou8RQDlVzr-diC28M3RUmUNao
by: Bill Steinberger
last edit: 1/6/2017
*/

var ss = SpreadsheetApp.getActiveSpreadsheet();
var pasteSheet = ss.getSheetByName("Shift Tracker");

function copyData() {
  var start = new Date('01/01/2017'); //set this date to whenever you want to start looking for data
  var newYear = new Date('01/01/2017'); //set this date to first day of the year you are in
  var now = new Date();

  var importSheet = ss.getSheetByName("Imported production data");
  if (now > newYear) {
    importSheet = ss.getSheetByName("2017 Imported production data");
  }

  var importData = importSheet.getDataRange().getValues();
  
  while(start < now){
    var targetDate = findDate(start);
    var prodData = {};
    //Logger.log(targetDate.month+' '+targetDate.day+' '+targetDate.year);
    for (var i=1; i<importData.length; i++) {
      var cell = importData[i][0];
      var testMonth = getCharsBefore(cell,'/');
      var testDay = getCharsAfter(cell,'/');
      if ((testMonth == targetDate.month) && (testDay == targetDate.day)) {
        prodData = {
          firstEast:importData[i][2],
          firstWest:importData[i][5],
          secondEast:importData[i+1][2],
          secondWest:importData[i+1][5],
          thirdEast:importData[i+2][2],
          thirdWest:importData[i+2][5]
        };
      }
    }
    pasteResults(targetDate,prodData);

    var newDate = start.setDate(start.getDate() + 1);
    start = new Date(newDate);
  }
}

function pasteResults(date,production) {
  var pasteData = pasteSheet.getDataRange().getValues();
  for (var i=1; i<pasteData.length; i++) {
    var cell = pasteData[i][0];
    if ((cell instanceof Date) && ((cell.getMonth()+1) == date.month) && (cell.getDate() == date.day) && (cell.getFullYear() == date.year)) {
      pasteSheet.getRange(i+1,2).setValue(production.firstEast);
      pasteSheet.getRange(i+1,3).setValue(production.firstWest);
      pasteSheet.getRange(i+1,4).setValue(production.secondEast);
      pasteSheet.getRange(i+1,5).setValue(production.secondWest);
      pasteSheet.getRange(i+1,6).setValue(production.thirdEast);
      pasteSheet.getRange(i+1,7).setValue(production.thirdWest);
    }
  }  
}

function findDate(date) {
  var yesterday = new Date(date.getTime() - (1 * (1000 * 60 * 60 * 24)));
  var yesterdayMonth = yesterday.getMonth()+1;
  var yesterdayDay = yesterday.getDate();
  var yesterdayYear = yesterday.getFullYear();
  return {
    month:yesterdayMonth,
    day:yesterdayDay,
    year:yesterdayYear
  };
}

function getCharsBefore(str, chr) {
    var index = str.indexOf(chr);
    if (index != -1) {
        return(str.substring(0, index));
    }
    return("");
}

function getCharsAfter(str, chr) {
    var index = str.indexOf(chr);
    if (index != -1) {
        return(str.substring((index+1), str.length));
    }
    return("");
}

function gotoToday() {
  var data = pasteSheet.getDataRange().getValues();
  var now = new Date();
  for (var i=1; i<data.length; i++) {
    var cell = data[i][0];
    if ((cell instanceof Date) && ((cell.getMonth()+1) == now.getMonth()+1) && (cell.getDate() == now.getDate()) && (cell.getFullYear() == now.getFullYear())) {
      var range = pasteSheet.getRange((i+1),1);
      pasteSheet.setActiveSelection(range);
      break;
    }
  }  
}

