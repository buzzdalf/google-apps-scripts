/*
** Set triggers up, tied to runAll for the following conditions: onEdit, onFormSubmit, time driven (midnight - 1 AM)
**  Setup another trigger to onOpen to run on open.
** Tied to: https://docs.google.com/spreadsheets/d/1Q5wPq6cfALLFh2sJY667DjSU3xrUaQgBrtm1WCk_L2Y
*/

var form = FormApp.openById('1cu9D6AEBpHYp3z8j1vt5plTD8GGJP4VvC5aev6c1HSQ');
var ss = SpreadsheetApp.openById('1Q5wPq6cfALLFh2sJY667DjSU3xrUaQgBrtm1WCk_L2Y');

function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('Refresh')
  .addItem('Update Area Listing', 'buildAreas')
  .addItem('Update Forms Data', 'runAll')
  .addToUi();  
}

function runAll() {
  refreshDates();
  buildList();
}

function buildList() {
  var masterSheet = ss.getSheetByName('master');
  var masterData = masterSheet.getDataRange().getValues();
  var allPages = form.getItems(FormApp.ItemType.PAGE_BREAK);
  var allLists = form.getItems(FormApp.ItemType.LIST);
//  var pages = getPages(allPages);
  var list = getLists(allLists);
  var listNames = addNames(masterData, list);
  
  updateForm(allLists,listNames);
}

function getLists(allLists) {
  var list = [];
  for (var i=0;i<allLists.length;i++) {
    var listTitle = allLists[i].getTitle();
    var listInfo = {
      title: listTitle,
      id: allLists[i].getId(),
      names: [],
      area: listTitle.substring(0, listTitle.indexOf(" Person being audited"))
    };
    list.push(listInfo);
//    Logger.log(listInfo.title+' '+listInfo.id+' '+listInfo.names+' '+listInfo.area);
  }
//  Logger.log(list);
  return list;
}

function addNames(masterData, list) {
  var testDate = getStartTime().start;
  //  Logger.log(testDate);
  for (var i=0;i<masterData.length;i++) {
    var lastDate = masterData[i][4];  //use date column-1 here
    //    Logger.log(lastDate);
    if (lastDate < testDate) {
      //        Logger.log('match');
      for (var j=0;j<list.length;j++) {
//        Logger.log('masterdata0:'+masterData[i][0]+' area:'+list[j].area+' masterdata1:'+masterData[i][1]);
        if (masterData[i][0] == list[j].area && masterData[i][1]) {  //ues area column-1 & name column-1 here
//          Logger.log(list[j].area+' '+masterData[i][1]);
          list[j].names.push(masterData[i][1]); //use name column-1 here
        }
//        Logger.log(list[j]);
      }
    }
  }
//  Logger.log(list);
  return list;
}

function updateForm(allLists,list) {
  for (var i=0;i<allLists.length;i++) {
    //    try {
    var myList = allLists[i].asListItem();
//    Logger.log(allLists[i]+' '+list[i].names);
    if (list[i].area) {
      Logger.log(list[i]);
      if (list[i].names.length>0) {
        myList.setChoiceValues(list[i].names);
      } else {
        myList.setChoiceValues(['']);
      }
    }
    //    } catch (e) {
    //      logError(e);
    //    }
  }
}

// this function take's a date and returns the month / day / year
function getDate(now) {
  if (now instanceof Date) {
    var day = now.getDate();
    var month = now.getMonth()+1;
    var year = now.getFullYear();
    var fullDate = month + "/" + day + "/" + year; 
    return fullDate;
  }
}

function getStartTime() {
  var today = new Date();
  var i = 365; //set this to the number of days you want to look back
  var yesterday = new Date(today.getTime() - i * 24 * 60 * 60 * 1000);
  var day = yesterday.getDate();
  var month = yesterday.getMonth()+1;
  var year = yesterday.getFullYear();
  var start = yesterday;
  var startString = yesterday.toISOString();
  return {
    day:day,
    month:month,
    year:year,
    start:start,
    startString:startString
  };
}

function refreshDates() {
  var masterSheet = ss.getSheetByName('master');
  var masterData = masterSheet.getDataRange().getValues();
  
  for (var i=1;i<masterData.length;i++) {
    var masterName = masterData[i][1]; //use name column-1 here
    var masterDate = masterData[i][4]; //use date column-1 here
    if (masterName) {
      var lastDate = checkForm(ss,masterName);
      if(lastDate) {
        masterSheet.getRange(i+1,5).setValue(lastDate); //use date column here
      }
    }
  }
}

function checkForm(ss,masterName) {
  var formSheet = ss.getSheetByName('Form Responses 1');
  var formData = formSheet.getDataRange().getValues();
  
  for (var j=0;j<formData.length;j++) {
    for (var k=0;k<formData[j].length;k++) {
      var testCell = formData[j][k];
      if (testCell == masterName) {
        var date = formData[j][0];
//        Logger.log('match'+' '+masterName+' '+date);
      }
    }
  }
  return date;
}

function buildAreas() {
  var allLists = form.getItems(FormApp.ItemType.LIST);
  var list = getLists(allLists, form);
  var areaId = findAreaList(list);
  var areaList = form.getItemById(areaId).asListItem();
  var choices = areaList.getChoices();
  
  var choiceList = [];
  for (var i=0;i<choices.length;i++) {
    var test = choices[i].getValue();
    choiceList[i] = [];
    choiceList[i].push(test);
  }
  fillAreas(choiceList);
}

function fillAreas(areas) {
  var areaSheet = ss.getSheetByName('areas');
  var lastRow = areaSheet.getLastRow();
  areaSheet.getRange(2,1,lastRow,1).clear();
  areaSheet.getRange(2,1,areas.length,1).setValues(areas);
}

function findAreaList(list) {
  for (var i=0;i<list.length;i++) {
//    Logger.log(list[i].title);
    if (list[i].title == 'Area Audit Conducted') {
      return list[i].id;
    }
  }
  return;
}

function getPages(allPages) {
  var pages = [];
  for (var i=0;i<allPages.length;i++) {
    var pageInfo = {
      title: allPages[i].getTitle(),
      id: allPages[i].getId()
    };
    pages.push(pageInfo);
  }
  return pages;
}


// the functions below are for error handling
function logError(e) {
  var url = activeUrl();
  var email = "william_f_steinberger@whirlpool.com";
  
//  MailApp.sendEmail(email, "Error report", 
//                    "\r\nMessage: " + url
//                    + "\r\nMessage: " + e.message
//                    + "\r\nFile: " + e.fileName
//                    + "\r\nLine: " + e.lineNumber);
  
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
