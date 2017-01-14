/* 
*  Script to automatically pick a random item from a list
*  also increments the item in the list to prevent repeats
* tied to: https://docs.google.com/spreadsheets/d/1eAeBuKo1iJwYQf3p9A0qbysIy4oUNspPWURDWUQnwgc
*  by: Bill Steinberger, please contact me with any issues or questions
*/

var usedColumn = 2;
var countColumn = 4;
var totalColumns = 4;
var startRow = 10;
var resultRow = 7;
var sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("picker");
var lastRow = sheet1.getLastRow();
var Avals = sheet1.getRange(startRow,1,lastRow,1).getValues();
var totalRows = Avals.filter(String).length;

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
//  Logger.log("Random Row:"+randRow);
  increment(randRow);
}

function randPick() {
  var randNumber = Math.floor((Math.random() * totalRows) + 1);
  var randRow = randNumber+startRow-1;
  return randRow;
}

function increment(row) {
  var count = sheet1.getRange(row,4).getValue();
  if (isNaN(parseFloat(count)) || !isFinite(count)) {
    count = 0;
  }
  count++;
  sheet1.getRange(row,4).setValue(count);
//  Logger.log("newCount:"+count);
}

function keepRunning() {
  for (var i=0;i<1000;i++) {
    picker();
  }
}
