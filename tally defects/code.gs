var ss = SpreadsheetApp.getActiveSpreadsheet();
var outputSheet = ss.getSheetByName('By Date'); //this is the name of the output sheet tab

//add functions for each button you want to incrment on
function increment1() {
  SpreadsheetApp.getActiveSheet().getRange('A1').setValue(SpreadsheetApp.getActiveSheet().getRange('A1').getValue() + 1);
}
function increment2() {
  SpreadsheetApp.getActiveSheet().getRange('c1').setValue(SpreadsheetApp.getActiveSheet().getRange('c1').getValue() + 1);
}
function increment3() {
  SpreadsheetApp.getActiveSheet().getRange('e1').setValue(SpreadsheetApp.getActiveSheet().getRange('e1').getValue() + 1);
}


//this function takes the counts and moves them to the total tab
function total() {
  var outputData = outputSheet.getDataRange().getValues();
  var inputSheet = ss.getSheetByName('data entry'); // this is the name of the input sheet tab
  var emptyRow = outputData.length + 1;
  var count = [];
  var title = [];
  
  for (var i=1;i<20;i+=6) {  //this is rows
    for (var j=1;j<8;j+=2) {  //this is columns
      var label = inputSheet.getRange(i+1,j).getValue();
      var value = inputSheet.getRange(i,j).getValue();
      if (label != "") {
        count.push(value);
        title.push(label);
      }
      inputSheet.getRange(i,j).setValue('0'); //resets the counter to 0
    }
  }
  saveOutput(emptyRow,count,title);
}

function saveOutput(row,output,title) {
  var today = new Date();
  var dateFormat = "MM/dd/yyyy hh:mm a" ;
  var date = Utilities.formatDate(today, Session.getScriptTimeZone(), dateFormat);
  
  outputSheet.getRange(row, 1).setValue(date); //save the current date to the output tab
  for (var j=0;j<output.length;j++) {
    outputSheet.getRange(row,j+2).setValue(output[j]);
    outputSheet.getRange(1,j+2).setValue(title[j]);
  }
}