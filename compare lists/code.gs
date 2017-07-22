/*
*  Tied to: https://docs.google.com/spreadsheets/d/1eCSUI0N_WJ8xgxF_Tl9_PEq5kbTKs14volcW9qEteHs
*/

var ss = SpreadsheetApp.openById('1eCSUI0N_WJ8xgxF_Tl9_PEq5kbTKs14volcW9qEteHs');

function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('Compare Lists')
  .addItem('Find items without a match', 'getData')
  .addToUi();  
}


function getData() {
  var totalSheet = ss.getSheetByName('total list');
  var totalData = totalSheet.getDataRange().getValues();
  var completedSheet = ss.getSheetByName('completed units');
  var completedData = completedSheet.getDataRange().getValues();
  
  var outputList = [];
  var j = 0;
  for (var i=1;i<totalData.length;i++) {
    var temp = totalData[i];
    var foundItem = compare(completedData,temp);
    if (foundItem) {
      outputList[j] = [];
      outputList[j].push(foundItem);
      j+=1;
    }
  }
  pasteNotFound(outputList);  
}

function compare(data,inputRow) {
  for (var i=1;i<data.length;i++) {
    if (data[i][0] == inputRow) {
      return;
    }
  }
  return inputRow;
}

function pasteNotFound(list) {
  var pasteSheet = ss.getSheetByName('not found');
  var lastRow = pasteSheet.getLastRow();
  pasteSheet.getRange(2,1,lastRow,1).clear();
  pasteSheet.getRange(2,1,list.length,1).setValues(list);
}