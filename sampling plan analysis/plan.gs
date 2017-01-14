/**
* Script to generate sampling plans, data entry forms and
* analyze the data as either an MSE or a COV.
* Implemented: All analysis, control charts, dot frequency
* no more planned features to be implemented
* tied to: https://docs.google.com/spreadsheets/d/1snFMbCTibvk7ZeS9oZTyLvs8dlh24qcVpcqqwV4oAG0
* Provide comments or issues to: Bill Steinberger
*/


//function onOpen() {
//  showSidebar();
//}

function setupSheets() {

  var ss = ssName();
  var inputSheet = inputSheetName();
  var rowCount = 1; //total number of factors
  var colCount = 1; //total number of data points the user will enter per the plan
  var rowLabel = []; //stores the name of factor label
  var colLabel = []; //a 2D array, storing each individual value on the input sheet
  var numCols = [];  //store the number of levels for each factor
  var rowColor = []; //stores the colors used for each factor on the input sheet
  var isNested = []; //boolean array, true if factor is nested
  var confCount = [];
  var inputData = [];
  var factor = [];
  var sampleSize = [];
  
//  popUp("setupSheets"); //Launch a dialog to let user know the script is running, so please wait.
  collectInputs();

  // get Sampling plan inputs
  function collectInputs() {
    var rowExist = true; //boolean to determine if there is a factor in this row
    var confTest = 0;
    var i = 2;
    inputData = inputSheet.getRange(2,1,11,11).getValues();
    for (var j=0;j<11;j++) {
      if (inputData[j][0] == "") {
        inputData.splice(j,(11-j));
        break;
      }
      inputData[j] = inputData[j].filter(function(e){return e});
      rowColor[j] = inputSheet.getRange(j+2,1).getFontColor();
      rowLabel[j] = inputData[j][0];
      colLabel[j] = [];
      for (var k=0;k<(inputData[j].length-1);k++) {
        colLabel[j][k] = inputData[j][k+1];
        if (colLabel[j][k] == "n" || colLabel[j][k] == "N") {
          isNested[j] = true;
        }
      }
      if (isNested[j]) {
        numCols[j] = k-1;
        confTest++;
        confCount[j] = confTest;
      } else {
        numCols[j] = k;
        confTest = 0;
      }
        colCount = colCount * numCols[j];
      }
    rowCount = j - 1;
    samplingPlan();
  }
  
  //populate Sampling Plan sheet
  function samplingPlan() {
    popUp("samplingPlan");
    var defaultName = "Sampling Tree";
    var treeSheetName = inputSheet.getRange(30,2).getValue();
    var planSheet = buildSheet(treeSheetName,defaultName); //name of sheet for the sampling plan
    var sheetName = planSheet.getName();
    inputSheet.getRange(30,2).setValue(sheetName); //set the sampling plan label to the result of duplicate test
    var outputRow = rowCount * 2 + 2; //last row of sampling plan (bottom of tree)
    var step = 1;
    if (colCount > 25) {
      planSheet.insertColumns(1,(colCount-25));
    }
    for (var i = rowCount;i>=0; i--) {
      factor[i] = [];
      factor[i].push(rowLabel[i]);
      var nestedCount = 1;
      var crossedCount = 0;
      var confLayer = (i-confCount[i]);
      for (var j = 2;j<=(colCount+1); j+=step) {
        var tempvar;
        if (isNested[i]) {
            if (i == rowCount || numCols[i] != 1) {
              tempvar = nestedCount;
            } else {
              tempvar = colLabel[confLayer][crossedCount];
            }
        } else {
          tempvar = colLabel[i][crossedCount];
        }
        factor[i].push(tempvar);
        if (i == rowCount) {
          planSheet.setColumnWidth(j, 25);
        }
        nestedCount++;
        if (crossedCount <= numCols[i]-2 || (isNested[i] && crossedCount <= numCols[confLayer]-2)) {
          crossedCount++;
        } else {
          crossedCount = 0;
        }
      }
      var tempList = [fillSpaces(step,factor[i])];
      var formatRange = planSheet.getRange(outputRow,1,1,tempList[0].length);
      formatRange.setValues(tempList);
      formatCell(formatRange,rowColor[i]);
      var labelRange = planSheet.getRange(outputRow-1,1,2,1);
      labelRange.merge();
      labelRange.setVerticalAlignment("middle");
      labelRange.setHorizontalAlignment("right");
      for (var j=2;j <=(colCount+1);j+=step) {
        planSheet.getRange(outputRow-1,j,2,step).merge();
      }
      outputRow-=2;
      sampleSize[i] = step;
      step = step * numCols[i];
    }
    planSheet.autoResizeColumn(1);
    populateData();
    closePopup(); //close the dialog, the script is done
  }

  // populate Data entry sheet
  function populateData() {
    popUp("populateData");
    var defaultName = "DataTable";
    var dataSheetName = inputSheet.getRange(31,2).getValue();
    var dataSheet = buildSheet(dataSheetName,defaultName); //name of sheet for data entry
    var sheetName = dataSheet.getName();
    inputSheet.getRange(31,2).setValue(sheetName); //set the data sheet label to the result of duplicate test
    var tempList = []
    var outputCol = rowCount + 1;
    for (var i = rowCount;i>=0; i--) {
      tempList[i] = fillData(factor[i],sampleSize[i]);
      var formatRange = dataSheet.getRange(1,outputCol,tempList[i].length,1);
      formatRange.setValues(tempList[i]);
      dataSheet.autoResizeColumn(outputCol);
      formatCell(formatRange,rowColor[i]);
      outputCol--;
    }
    dataColumn(dataSheet);
    deleteEmpties(dataSheet);
    closePopup(); //close the dialog, the script is done
  }

  // setup the data entry column
  function dataColumn(dataSheet) {
    var dataCol = rowCount+2;
    dataSheet.getRange(1,dataCol).setValue("Data Entry");
    dataSheet.autoResizeColumn(dataCol);
    dataSheet.setFrozenRows(1);
    var blankCell = dataSheet.getRange(2,dataCol,colCount,1);
    blankCell.setBackgroundColor("#FFFF99");
    blankCell.setBorder(true, false, true, false, false, true);
    var range = dataSheet.getRange(2, dataCol);
    dataSheet.setActiveSelection(range);
    return;
  }

  //delete extra columns & rows
  function deleteEmpties(dataSheet) {
    var totalRows = dataSheet.getMaxRows();
    dataSheet.deleteColumns((rowCount+3),(26-(rowCount+2)));
//    dataSheet.deleteRows((colCount+3),(totalRows-colCount-2));
    return;  
  }
}

// format the sampling tree and data table cells
function formatCell(cell,color) {
  cell.setHorizontalAlignment("center");
  cell.setVerticalAlignment("middle");
  cell.setFontColor(color);
  cell.setBorder(true, true, true, true, true, true, color, null);
}

// get the spreadsheet name and the input sheet name
function ssName() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return (ss);
}
function inputSheetName() {
  var ss = ssName();
  var inputSheet = ss.getSheetByName("input"); //name of the sheet to pull the inputs from
  return (inputSheet);
}

// Test to see if sheet names from the input tab already exist & if they do, increment them
function buildSheet(sheetName,defaultName) {
  var ss = ssName();
  var sheets = ss.getSheets();
  var sheetCounter = 0; //counter to increment sheet names if ok = false
  if (sheetName == "") {
    sheetName = defaultName;
  }
  var returnSheet = sheetName;
  for (var i=0;i<50;i++) {
    var ok = true; //boolean test of whether a sheet name is available (false = name taken)
    for ( var j=0; j<sheets.length;j++ ) {
      if (sheets[j].getName() == returnSheet) {
        ok = false;
        sheetCounter++;
      } 
    }
    if (ok) {
      var newSheet = ss.insertSheet(returnSheet);
      return(newSheet);
    } else {
      returnSheet = defaultName + sheetCounter; //if the name is taken use the standard name and increment
    }
  }
}

// Fill in the blank spaces for each factor on the Sampling Plan tab
function fillSpaces(sampleSize,variable) {
  var returnVariable = [];
  for (var i=0;i<variable.length;i++) {
    if (i==0) {
      returnVariable.push(variable[i]);
    } else {
      for (var h=1;h<=sampleSize;h++) {
        var tempVar = [];
        if (h == 1) {
          if (variable.constructor === Array) {
            tempVar.push(variable[i]);
          } else {
            tempVar.push(variable);
          }
        } else {
          tempVar.push("");
        }
        returnVariable.push([tempVar]);
      }
    }
  }
  return returnVariable;
}

//Fill the subgroup data into an array for the data entry sheet
function fillData(factor,size) {
  var returnVariable = [];
  returnVariable.push([factor[0]]);
  for (var j=1;j<factor.length;j++) {
    for (var h=1;h<=size;h++) {
      returnVariable.push([factor[j]]);
    }
  }
  return returnVariable;
}

// Show the sidebar menu
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Sampling Plan Analyzer')
  SpreadsheetApp.getUi()
      .showSidebar(html);
}


