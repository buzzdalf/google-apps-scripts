// script still in development, lacks ability to email/print reports.
// tied to: https://docs.google.com/spreadsheets/d/175PdflO3MNEDW3MVkqgl-wDy5RV2PA2aok1BUsyiN8E
// by: Bill Steinberger

// @param {Object} e The event parameter for a simple onOpen trigger.
function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('Setup Reports')
  .addItem('Instructions', 'runInstructions')
  .addItem('Shipment Weight and Measure', 'runCountSheet')
  .addItem('Staging Label', 'runStageLabel')
  .addItem('Load Verification Form', 'runLoadVerification')
  .addSeparator()
  .addItem('Print All', 'printStuff')
  .addToUi();
}

function runInstructions() {
  instructions();
  var sheetName = instructions();
  savePDF(sheetName);
}

function runCountSheet() {
  countSheet();
  var sheetName = countSheet();
  savePDF(sheetName);
}

function runStageLabel() {
  stageLabel();
  var sheetName = stageLabel();
  savePDF(sheetName);
}

function runLoadVerification() {
  loadVerification();
  var sheetName = loadVerification();
  savePDF(sheetName);
}

function instructions() {
  var data = getData();
  var outSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Instructions");

  outSheet.getRange(1,9).setValue(data.date);
  outSheet.getRange(2,3).setValue(data.inNum);
  outSheet.getRange(4, 3).setValue(data.store);
  outSheet.getRange(6, 6).setValue(data.wave);
  outSheet.getRange(11, 3).setValue(data.weight);
  outSheet.getRange(12, 3).setValue(data.mode);
  outSheet.getRange(17, 2).setValue(data.address);
  outSheet.getRange(23, 2).setValue(data.fwd);
  outSheet.getRange(13, 3).setValue(data.terms);
  outSheet.getRange(1, 14).setValue(data.rep);
  outSheet.getRange(7, 2).setValue(data.po);
  outSheet.getRange(5, 10).setValue(data.fo);
  outSheet.getRange(22, 11).setValue(data.update);
  outSheet.getRange(6, 8).setValue(data.ltrCode);
  outSheet.getRange(6, 8).setBackground(data.color);
  outSheet.getRange(23, 10).setValue(data.notes);
  
  return outSheet;
}

function countSheet() {
  var data = getData();
  var outSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Shipment Weight and Measure");
  
  outSheet.getRange(4, 2).setValue(data.inNum);
  outSheet.getRange(5, 2).setValue(data.wave);
  outSheet.getRange(6, 2).setValue(data.mode);
  outSheet.getRange(1, 2).setValue(data.address);
  outSheet.getRange(3, 6).setValue(data.ltrCode);
  outSheet.getRange(3, 6).setBackground(data.color);
  
  return outSheet;

}


function stageLabel() {
  var data = getData();
  var outSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Staging Label");

  outSheet.getRange(2,3).setValue(data.inNum);
  outSheet.getRange(1, 7).setValue(data.wave);
  outSheet.getRange(5, 3).setValue(data.weight);
  outSheet.getRange(4, 7).setValue(data.mode);
  outSheet.getRange(8, 7).setValue(data.address);
  outSheet.getRange(8, 2).setValue(data.fwd);
  outSheet.getRange(1, 8).setValue(data.ltrCode);
  outSheet.getRange(1, 7).setBackground(data.color);
  outSheet.getRange(1, 8).setBackground(data.color);
  
  return outSheet;

}


function loadVerification() {
  var data = getData();
  var outSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Load Verification Form");

  outSheet.getRange(3, 13).setValue(data.wave);
  outSheet.getRange(3, 14).setValue(data.ltrCode);
  outSheet.getRange(3, 13).setBackground(data.color);
  outSheet.getRange(3, 14).setBackground(data.color);
  
  return outSheet;

}


function printStuff() {
  runInstructions();
  runCountSheet();
  runStageLabel();
  runLoadVerification();
  
}


function getData() {
  var dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");
  var currentRow = dataSheet.getActiveCell().getRowIndex();
  
  var inData = {
    date: dataSheet.getRange(currentRow,1).getValue(),
    inNum: dataSheet.getRange(currentRow,8).getValue(),
    wave: dataSheet.getRange(currentRow,13).getValue(),
    address: dataSheet.getRange(currentRow,5).getValue(),
    fwd: dataSheet.getRange(currentRow,6).getValue(),
    terms: dataSheet.getRange(currentRow,12).getValue(),
    rep: dataSheet.getRange(currentRow,2).getValue(),
    po: dataSheet.getRange(currentRow,4).getValue(),
    update: dataSheet.getRange(currentRow,14).getValue(),
    fo: dataSheet.getRange(currentRow,7).getValue(),
    notes: dataSheet.getRange(currentRow,16).getValue(),
    ltrCode: dataSheet.getRange(currentRow,15).getValue(),
    store: dataSheet.getRange(currentRow,3).getValue(),
    color: dataSheet.getRange(currentRow,15).getBackground(),
    weight: dataSheet.getRange(currentRow,9).getValue(),
    mode: dataSheet.getRange(currentRow,10).getValue()
  }; 
  
  return inData;
  
}

function createFolder(ss) {
  var driveFolder = "print queue"; //folder name
  var parents = DriveApp.getFileById(ss.getId()).getParents();
  if (parents.hasNext()) {
    var parentFolder = parents.next();
  }
  else {
    parentFolder = DriveApp.getRootFolder();
  }
  var folders = DriveApp.getFoldersByName(driveFolder); 
  var folder = folders.hasNext() ? folders.next() : parentFolder.createFolder(driveFolder); //using parentFolder creates a sub-folder inside the current folder
  //var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(driveFolder); //using DriveApp will create a sub-folder in the root
  return folder;
}

function savePDF (sheet) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var url = ss.getUrl().replace(/edit$/,'');

  var folder = createFolder(ss);

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
  var blob = response.getBlob().setName(sheet.getName() + '.pdf');
  folder.createFile(blob);
}

/**
 * Dummy function for API authorization only.
 * From: http://stackoverflow.com/a/37172203/1677912
 */
function forAuth_() {
  DriveApp.getFileById("Just for authorization"); // https://code.google.com/p/google-apps-script-issues/issues/detail?id=3579#c36
}