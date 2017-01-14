/**
* Script to search the sheet for specific audits
* to help users find item to close.
* tied to: https://docs.google.com/spreadsheets/d/1Oh4z8oHqlpbokIz2HftcLMAzK6_yFJmWbL3jBbu2Byk
* Provide comments or issues to: Bill Steinberger
* Version 1.5 05/08/2016
* search working for area and auditor
* cleaned up code ,eliminating redunancies
* added a search by business unit
* changed columns searched from number to labels in case they move around
* testing Search by responsible party - currently working for corrective action 1 only
*/

var SIDEBAR_TITLE = 'Find Your Audits';
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data of Form Responses");
var data = sheet.getDataRange().getValues();
var lastRow = sheet.getLastRow();

//setup all column labels for searches
var areaText = 'Area';
var auditorText = 'Username';
var responsible1Text = 'Corrective Action 1 Responsible Party';
var responsible2Text = 'Corrective Action 2 Responsible Party';
var responsible3Text = 'Corrective Action 3 Responsible Party';
var buText = 'Business Unit';
var closedText = 'All Corrective Actions Closed?';
var idText = 'Timestamp';

// @param {Object} e The event parameter for a simple onOpen trigger.
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Find Audits')
      .addItem('Search Audits by Area', 'searchArea')
      .addItem('Search Audits by Auditor', 'searchAuditor')
      .addItem('Search Audits by Responsible Party', 'searchResponsible')
      .addItem('Search Audits by Business Unit', 'searchBU')
      .addToUi();
}

function showSidebar(menu) {
  var ui = HtmlService.createTemplateFromFile(menu)
      .evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle(SIDEBAR_TITLE);
  SpreadsheetApp.getUi().showSidebar(ui);
}

function searchArea() {
  var myCol = findColumn(areaText);
  testProps(myCol);
}

function searchAuditor() {
  var myCol = findColumn(auditorText);
  testProps(myCol);
}

function searchResponsible() {
//  var myCol = {
//    1: findColumn(responsible1Text),
//    2: findColumn(responsible2Text),
//    3: findColumn(responsible3Text)
//  };
  var myCol = findColumn(responsible1Text);
  testProps(myCol);
}

function searchBU() {
  var myCol = findColumn(buText);
  testProps(myCol);
}

function testProps(setCol) {
  do {
  setProps(setCol);
  var getCol = getProps();
  }
  while (getCol != setCol);
  
  showSidebar('Search');
}

function setProps(myCol) {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.deleteAllProperties();
//  for (i=1;i<=3;i++) {
//    userProperties.setProperty('col'+i, myCol.i);
//  }
  userProperties.setProperty('col', myCol);
}

function getProps() {
  var userProperties = PropertiesService.getUserProperties();
  var tempCol = parseInt(userProperties.getProperty('col'));
  return tempCol;
}

function getColumn(data,col) {
  return data.map(function(value,index) { 
    return value[col]; 
  }); 
}

function findUnique() {
  var col = getProps();
  var colData = getColumn(data,col);
  colData.shift();  //remove the header row from the array
  var newdata = colData.filter(function(elem, pos) {  //use filter to find the unique values
    return colData.indexOf(elem) == pos;
  }); 
  newdata.sort();
  return newdata;
}

function findAudits(selected,all) {
  var col = getProps();
  var listRows = [];
  var openCol = findColumn(closedText);
  var idCol = findColumn(idText);
  for (var i in data) {
    if (data[i][col] == selected) {
      var value = data[i][idCol];
      value = Utilities.formatDate(value, Session.getScriptTimeZone() , "M/d/yyyy");
      if (all || data[i][openCol] == "No") {
        listRows.push({record:i,timestamp:value});  //actual row number is i+1
      }
    }
  }
  return listRows;
}

function findColumn(criteria) {
  for (i=1;i<=data[0].length;i++) {
    if (data[0][i] == criteria) {
      return i;
    }
  }
  return;
}

function gotoRow(selectedRow) {
  var range = sheet.getRange(selectedRow,41);
  Logger.log("selected row: "+selectedRow);
  sheet.setActiveSelection(range);
}
