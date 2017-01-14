/*
*  Code to calculate the samples for the Sampling theory exercises (bead box)
*  Version 1.0; 7/14/2016
* tied to: https://docs.google.com/spreadsheets/d/1f6-FMUcfgMyrUB_HYR1ETRk1HmRjlACMbb2TPTYFLT0
*  Any questions contact: Bill Steinberger
*/

var setupSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("setup");

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Re-Calc Menu')
      .addItem('Run Part 1', 'part1')
      .addItem('Run Part 2', 'part2')
      .addItem('Run Part 3', 'part3')
      .addSeparator()
      .addToUi();
}

function part1() {
  var part1Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Part 1");
  SpreadsheetApp.setActiveSheet(part1Sheet);
  var pGood = setupSheet.getRange(1,3).getValue();
  // guide to run the functions: getColors(#Rows,#Columns)
  var range5 = getColors(5,1,pGood);
  var range10 = getColors(5,2,pGood);
  var range50 = getColors(10,5,pGood);
  var range100 = getColors(20,5,pGood);
  part1Sheet.getRange(8,2,5,1).setValues(range5);
  part1Sheet.getRange(8,5,5,2).setValues(range10);
  part1Sheet.getRange(8,8,10,5).setValues(range50);
  part1Sheet.getRange(8,14,20,5).setValues(range100);
}

function part2() {
  var part2Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Part 2");
  SpreadsheetApp.setActiveSheet(part2Sheet);
  var origGood = setupSheet.getRange(4,3).getValue();
  var newGood = setupSheet.getRange(7,3).getValue();
  // guide to run the functions: getColors(#Rows,#Columns)
  var rangeOld = getColors(10,1,origGood);
  var rangeNew = getColors(10,1,newGood);
  part2Sheet.getRange(7,2,10,1).setValues(rangeOld);
  part2Sheet.getRange(7,5,10,1).setValues(rangeNew);
}

function getColors(k,m,percentGood) {
  var colors = [];
  for (var i=0;i<k;i++) {
    colors[i] = [];
    for (var j=0;j<m;j++) {
      var rand = Math.random();
      //    Logger.log(rand);
      if (rand > (1-percentGood)) {
        colors[i][j] = "White";
      } else {
        colors[i][j] = "Red";
      }
    }
  }
  Logger.log(colors);
  return colors;
}

function part3() {
  var part3Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Part 3");
  SpreadsheetApp.setActiveSheet(part3Sheet);
  var allStats = setupSheet.getRange(11,4,6,2).getValues();
  var original = getValues(allStats[0]);
  var newA = getValues(allStats[1]);
  var newB = getValues(allStats[2]);
  var newC = getValues(allStats[3]);
  var newD = getValues(allStats[4]);
  var newE = getValues(allStats[5]);
  part3Sheet.getRange(6,2,10,1).setValues(original);
  part3Sheet.getRange(6,5,10,1).setValues(newA);
  part3Sheet.getRange(6,8,10,1).setValues(newB);
  part3Sheet.getRange(6,11,10,1).setValues(newC);
  part3Sheet.getRange(6,14,10,1).setValues(newD);
  part3Sheet.getRange(6,17,10,1).setValues(newE);
}

function getValues(stats) {
  var results = [];
  var mu = stats[0];
  var sigma = stats[1];
  for (var i=0;i<10;i++) {
    results[i] = [];
    var p = Math.random();
    results[i][0] = normsInv(p, mu, sigma);
  }
  Logger.log(results);
  return results;
}
