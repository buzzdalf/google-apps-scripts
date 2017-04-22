/* 
*  Script to a list from Column A in a random order and put the results in column B
*  tied to: https://docs.google.com/spreadsheets/d/1RxgzJj3hNJsHlL_4iNe25VxahH0qkUANOheJG8DhSdE
*  by: Bill Steinberger, please contact me with any issues or questions
*/

function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('Refresh List')
  .addItem('Randomize', 'randomizer')
  .addToUi();
}

function randomizer() {
  var startRow = 3;
  var resultColumn = 2;
  var sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("picker");
  sheet1.getRange("B3:B").clear();
  
  //first load everything from A3 down into a variable, determine how many rows have actual data and then get just the rows to last row with data
  var data = sheet1.getRange("A3:A").getValues();
  var totalRows = data.filter(String).length;  
  data = sheet1.getRange(3,1,totalRows,1).getValues();
  
  var output = [];
  var k = 0;
  output[k] = data[random(totalRows)];

  while (k < totalRows) {
    var randNumber = random(totalRows);
    var available = testRepeat(output,data[randNumber]);
    if (available) {
      output[k] = data[randNumber];
      k++;
    }
  }
  sheet1.getRange(startRow,resultColumn,totalRows,1).setValues(output);
}

function testRepeat(list,picked) {
  for (var j=0;j<list.length;j++) {
    if (picked === list[j]) {
      return false;
    }
  }
  return true;
}

function random(seed) {
  return (Math.floor(Math.random() * seed))
}
