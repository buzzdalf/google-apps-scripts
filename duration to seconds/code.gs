/*
* Script to read all data from an input sheet,
* find cells with time duration in min:sec.mmm format
* and convert them to sec.mmm format in an output sheet
* so the numbers can be used in calcuations
* This script is tied to: https://docs.google.com/spreadsheets/d/1y4qkOLllCSPcnT_wQaHzqDjO2_rB55PuL3mcKoSglEg
* By: Bill Steinberger Last rev: 01/18/2017
*/
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Update Averages')
      .addItem('Recalculate', 'convertTime')
      .addSeparator()
      .addToUi();
}

function convertTime() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet2');
  var data = sheet.getDataRange().getValues();
  var outSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('output');
  
  var colCount = data[0].length;
  var rowCount = data.length;
  for(var i=0;i<rowCount;i++) {
    for (var j=0;j<colCount;j++) {
      var contents = data[i][j].toString();
      var mins = parseInt(getCharsBefore(contents,':'));
      var secs = parseFloat(getCharsAfter(contents,':'));
      if(!isNaN(mins) && isFinite(mins) && !isNaN(secs) && isFinite(secs)) {
        var duration = mins*60+secs;
        data[i][j] = duration;
      }
    }
  }
  outSheet.getRange(1,1,rowCount,colCount).setValues(data);
  outSheet.getRange(3,1,rowCount,colCount).clearFormat();
}

function getCharsBefore(str, chr) {
  var index = str.indexOf(chr);
  if (index != -1) {
    return(str.substring(0, index));
  }
  return;
}

function getCharsAfter(str, chr) {
  var index = str.indexOf(chr);
  if (index != -1) {
    return(str.substring(index+1, str.length));
  }
  return;
}