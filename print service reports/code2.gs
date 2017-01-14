// script still in development, lacks ability to email/print reports.
// tied to: https://docs.google.com/spreadsheets/d/12DNpG2ImTqx8vaIJ1AzFU6GfcCXJ6v-0sHM9L_-16Yw
// by: Bill Steinberger

// @param {Object} e The event parameter for a simple onOpen trigger.
function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('Setup Reports')
  .addItem('Instructions', 'instructions')
  .addToUi();
}

function instructions() {
  var dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DATA");
  var allData = dataSheet.getDataRange().getValues();

  for (var i=2;i<=allData.length;i++) {
    var data = getData(dataSheet,i);
    var outSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Instruction sheet");
    
    outSheet.getRange(1, 2).setValue(data.inNum);
    outSheet.getRange(3, 2).setValue(data.po);
    outSheet.getRange(6, 2).setValue(data.carrier);
    outSheet.getRange(8, 2).setValue(data.pa);
    outSheet.getRange(7, 2).setValue(data.dat);
    outSheet.getRange(4, 6).setValue(data.fo);
    outSheet.getRange(13, 1).setValue(data.address);
    outSheet.getRange(17, 5).setValue(data.notes);
    outSheet.getRange(19, 1).setValue(data.shipTo);
    outSheet.getRange(1, 5).setValue(data.rdt);
    outSheet.getRange(1, 7).setValue(data.cRep);
    outSheet.getRange(15, 6).setValue(data.update);
    outSheet.getRange(9, 2).setValue(data.terms);
    outSheet.getRange(14, 7).setValue(data.store);
    var j = i-1;
    
    savePDF(outSheet,j);
  }  
}


function getData(dataSheet,currentRow) {
//  var currentRow = dataSheet.getActiveCell().getRowIndex();
  
  var inData = {
    inNum: dataSheet.getRange(currentRow,8).getValue(),
    po: dataSheet.getRange(currentRow,4).getValue(),
    carrier: dataSheet.getRange(currentRow,9).getValue(),
    pa: dataSheet.getRange(currentRow,13).getValue(),
    dat: dataSheet.getRange(currentRow,10).getValue(),
    fo: dataSheet.getRange(currentRow,7).getValue(),
    address: dataSheet.getRange(currentRow,5).getValue(),
    notes: dataSheet.getRange(currentRow,12).getValue(),
    shipTo: dataSheet.getRange(currentRow,6).getValue(),
    rdt: dataSheet.getRange(currentRow,1).getValue(),
    cRep: dataSheet.getRange(currentRow,2).getValue(),
    update: dataSheet.getRange(currentRow,11).getValue(),
    terms: dataSheet.getRange(currentRow,14).getValue(),
    store: dataSheet.getRange(currentRow,3).getValue()
  }; 
  
  return inData;
}


function createFolder(ss) {
  //creates the folder for storing all the charts if it doesn't already exist
  var driveFolder = "print queue";

  var parents = DriveApp.getFileById(ss.getId()).getParents();
  if (parents.hasNext()) {
    var parentFolder = parents.next();
  }
  else {
    parentFolder = DriveApp.getRootFolder();
  }

  var folders = DriveApp.getFoldersByName(driveFolder); 
  var folder = folders.hasNext() ? folders.next() : parentFolder.createFolder(driveFolder);
  return folder;
}

function savePDF (sheet, record) {
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
  var blob = response.getBlob().setName(sheet.getName() + ' - ' + record + '.pdf');
  folder.createFile(blob);
}

/**
 * Dummy function for API authorization only.
 * From: http://stackoverflow.com/a/37172203/1677912
 */
function forAuth_() {
  DriveApp.getFileById("Just for authorization"); // https://code.google.com/p/google-apps-script-issues/issues/detail?id=3579#c36
}