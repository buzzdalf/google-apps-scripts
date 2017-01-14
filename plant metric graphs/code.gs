/**
* Script to generate plant wide metric graphs
* Provide comments or issues to: Bill Steinberger
* In order to enable "update vs replace feature, you need to enable Drive API and
* use this developer console project: 127274555343
* Setup to Trigger buildCharts() every day at 3:00 AM if you want to automate
* tied to: https://docs.google.com/spreadsheets/d/1SOrneK8pr6T0M_Yhbrpr97sgfuBh9U2_oI7cqSjoLCA
* chart customization options available here: 
* https://developers.google.com/apps-script/reference/spreadsheet/embedded-chart-builder#methods
* https://developers.google.com/chart/interactive/docs/gallery/columnchart#configuration-options
*/

var dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");
var setupSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("setup");
var graphSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Charts");
var engageSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Engagement");
var hideText = "Hide Numbers";

function buildCharts() {
  var month = setupSheet.getRange(3,2,16,1);
  var metrics = [];
  var plotRow = 1;
  var colorRow = 7;
  var folder = createFolder();
  var apiEnabled = checkAPI(folder);
  var graphDoc = createDoc(folder);
  deleteCharts(graphSheet);
  
  for (var i=3;i<=87;i+=6) {
    var testData = setupSheet.getRange(3,i,16,1).getValues();
    var hasData = checkArray(testData);
    if (hasData) {
      var chartData = {
        green: setupSheet.getRange(3,i+1,16,1),
        red: setupSheet.getRange(3,i+2,16,1),
        prev: setupSheet.getRange(3,i+3,16,1),
        prevColor: dataSheet.getRange(colorRow,4).getBackground(),
        target: setupSheet.getRange(3,i+4,16,1),
        title: setupSheet.getRange(1,i).getValue(),
        hidden: dataSheet.getRange(colorRow,24).getValue(),
        row: plotRow
      };
      chart(chartData);
      plotRow+=34;
    }
    colorRow+=3;
  }
  copyGraph();
  engagement();
  
//Build a chart
  function chart(metric) {
    metrics.push(metric.title);
    var colorArray = [];
    var findRange = [];
    var textColor = 'black';
    var g = checkSeries(metric.green);
    if (g == 0) {
      colorArray.push('green');
      findRange.push(calcRange(metric.green.getValues()));
    }
    var rd = checkSeries(metric.red);
    if (rd == 0) {
      colorArray.push('red');
      findRange.push(calcRange(metric.red.getValues()));
    }
    var r = g + rd;
    var pr = checkSeries(metric.prev);
    if (pr == 0) {
      colorArray.push(metric.prevColor);
      findRange.push(calcRange(metric.prev.getValues()));
    }
    var p = r + pr;
    var tg = checkSeries(metric.target);
    if (tg == 0) {
      colorArray.push('black');
      findRange.push(calcRange(metric.target.getValues()));
    }
    var t = p + tg;
    if(metric.hidden == hideText) {
      textColor = 'white';
    }

    var chartRange = setBaseline(findRange);
    var chart1 = graphSheet.newChart()
    .addRange(month)
    .addRange(metric.green)
    .addRange(metric.red)
    .addRange(metric.prev)
    .addRange(metric.target)
    .setOption('colors',colorArray) 
    .setChartType(Charts.ChartType.COLUMN)
    .setPosition(metric.row,1,0,0)
    .setOption('width', 925)
    .setOption('height', 675)
    .setOption('legend', 'none')
    .setOption('title', metric.title)
    .setOption('titlePosition', 'out')
    .setOption('titleTextStyle.color','black')
    .setOption('titleTextStyle.fontSize','32')
    .setOption('titleTextStyle.bold','true')
    .setOption('titleTextStyle.italic','false')
    .setOption('vAxis.gridlines.count', 6)
    .setOption('vAxis.viewWindow.min', chartRange.min)
    .setOption('vAxis.minValue', chartRange.min)
    .setOption('vAxis.maxValue', chartRange.max)
    .setOption('vAxis.backgroundColor','black')
    .setOption('vAxis.textStyle.color',textColor)
    .setOption('hAxis.textStyle.bold','true')
    .setOption('annotations.textStyle.color',textColor)
    .setOption('annotations.textStyle.fontSize','16')
    .setOption('annotations.textStyle.auraColor','white')
    .setOption('annotations.textStyle.bold','true')
    .setOption('annotations.style','point')
    .setOption('annotations.stem.color','none')
    .setOption('annotations.highContrast', 'true')
    .setOption('annotations.alwaysOutside', 'true')
    .setOption('bar.groupWidth','80%')
    .setOption('isStacked','true')
    .setOption('series.0.dataLabel','value')
    .setOption('series.'+(1-g)+'.dataLabel','value')
    .setOption('series.'+(2-r)+'.dataLabel','value')
    .setOption('series.'+(3-p)+'.type','line')
    .setOption('series.'+(3-t)+'.role','annotation')
    .setOption('chartArea.left','100')
    .setOption('chartArea.width','775')
    .setOption('chartArea.top','100')
    .setOption('chartArea.height','500')
    .build();
    graphSheet.insertChart(chart1);
  }
  
  function copyGraph() {
  //copy the graphs out to the google drive for posting to sites, printing, etc
    var charts = graphSheet.getCharts();
    var chartBlobs = [];
    for (var i in charts) {
      if (!apiEnabled) {
        deleteFile(folder,metrics[i]);
      }
      chartBlobs[i]= charts[i].getAs('image/png').setName(metrics[i]);
      var fileExist = checkExist(folder,metrics[i]);
      if (fileExist) {
        var fileID = getFileId(folder,metrics[i]);
        replaceFile(chartBlobs[i], metrics[i], fileID);
      } else {
        folder.createFile(chartBlobs[i]);
      }
      updateDoc(i,chartBlobs[i]);
    }
  }
  
  function updateDoc(i,blob) {
    // add chart to the printable document
    var body = graphDoc.getBody();
     var image = body.insertImage(i,blob);
    image.setHeight(730)
         .setWidth(1000);
  }
  
  function checkArray(myArray){
  // check an array for data, determine if the array is blank, if it is blank, return false
    var count = 0;
    for(var i=0;i<myArray.length;i++){
       if(myArray[i] != "" && !isNaN(myArray[i]))   
      //if (!isNaN(myArray[i]) && isFinite(myArray[i]))
          return true;
    }
    return false;
  }
  
  function checkSeries(range) {
  //convert a range in sheet to an array to pass into the checkArray function
    var series = range.getValues();
    var test = checkArray(series);
    if (test) {
      return 0;
    } else return 1;
  }
  
  function calcRange(rangeArray){
  // find the min and max of each series
    var testArray = rangeArray.filter(function(number) {
      return (!isNaN(parseFloat(number)) && isFinite(number))
    });
    var rangeData = [];
    var rangeData = {
      max: testArray.sort(function(a,b){return b-a})[0][0],
      min: testArray.sort(function(a,b){return a-b})[0][0],
    };
    return rangeData;
  }

  function setBaseline(minArray){
  // calculate a custom starting Y value for the graph
    var minList = [];
    var maxList = [];
    for (var i=0;i<minArray.length;i++) {
      minList.push(minArray[i].min); 
      maxList.push(minArray[i].max);
    }
    var min = minList.sort(function(a,b){return a-b})[0];
    var max = maxList.sort(function(a,b){return b-a})[0];
    var range = max-min;
    var nice = {
      min: 0,
      max: max
    };
    if (range != 0) {
      var space = range/5;
      nice.min = Math.floor(min/space)*space-space;
      nice.max = max + space;
    }
    return nice;
  }
}

function deleteCharts(sheet) {
  //delete any existing charts before creating new ones
  var charts = sheet.getCharts();
  for (var i in charts) {
    sheet.removeChart(charts[i]);
  }
}

function createFolder() {
  //creates the folder for storing all the charts if it doesn't already exist
  var driveFolder = setupSheet.getRange(1,1).getValue() + ' ' + dataSheet.getRange(2,6).getValue();
  if (driveFolder == "" || driveFolder == null) {
    driveFolder = "PLANT METRIC DISPLAY";
  }
  var folders = DriveApp.getFoldersByName(driveFolder); 
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(driveFolder);
  return folder;
}

function getFileId(folder,fileName) {
  //gets the file ID for a given file
  var files = folder.getFilesByName(fileName);
  while (files.hasNext()) {
    var tempID = files.next().getId();
  }
  return tempID;
}

function deleteFile(folder,fileName) {
  //deleted existing files so new ones can be created, maintaining only the latest copy
  var files = folder.getFilesByName(fileName);
  while (files.hasNext()) {
    files.next().setTrashed(true);
  }
}

function checkAPI(folder) {
  //check to see if the Drive API is enabled so I know whether I can update files or have to replace them
  var name = 'Test';
  folder.createFile(name,'test file');
  var fileID = getFileId(folder,name); 
  try {
    Drive.Files.remove(fileID);
  } catch(e) {
    deleteFile(folder,name);
    return false;
  }
  return true;
}

function replaceFile(blob,fileName,fileID) {
  //updates the content of a file without creating a new file ID (requires Drive API)
  var file = {
    title: fileName,
    mimeType: 'image/png'
  };
  var myVar = Drive.Files.update(file, fileID, blob);
}

function checkExist(folder,fileName) {
  //checks if a file already exists
  var exist = true;
  var files = folder.getFilesByName(fileName);
  var testFile = files.hasNext() ? files.next() : exist = false;
  return exist;
}

function createDoc(name) {
  //prepares the printable document
  var fileExist = checkExist(name,name);
  //if the file already exists, simply clear the contents
  if (fileExist) {
    var fileID = getFileId(name,name);
    var doc = DocumentApp.openById(fileID);
    var body = doc.getBody();
    body.clear();
  } else {
    //if the file does not exist, create it and move it into the correct folder
    var doc = DocumentApp.create(name),
        docFile = DriveApp.getFileById( doc.getId() );
    name.addFile(docFile);
    DriveApp.getRootFolder().removeFile(docFile);
    var body = doc.getBody();
    var width = body.getPageHeight();
    var height = body.getPageWidth();
    body.setMarginBottom(30)
      .setMarginLeft(20)
      .setMarginRight(20)
      .setMarginTop(30)
      .setPageHeight(height)
      .setPageWidth(width);
  }
  return doc;
}

function engagement() {
  var labels = engageSheet.getRange(1,1,5,1);
  var prevSalary = engageSheet.getRange(1,2,5,1);
  var prevHourly = engageSheet.getRange(1,4,5,1);
  var salaryTarget = engageSheet.getRange(1,3,5,1);
  var hourlyTarget = engageSheet.getRange(1,5,5,1);
  var activities = engageSheet.getRange(8,1,3,2).getValues();
  var engageTitle = engageSheet.getRange(1,1).getValue();
  var salaryColor = engageSheet.getRange(1,1).getBackground();
  var folder = createFolder();
  var apiEnabled = checkAPI(folder);

  deleteCharts(engageSheet);
  engageChart();
  copyEngage();
  
  //Build engagement chart
  function engageChart() {
    var engageChart = engageSheet.newChart()
    .addRange(labels)
    .addRange(prevSalary)
    .addRange(salaryTarget)
    .addRange(prevHourly)
    .addRange(hourlyTarget)
    .setOption('colors', [salaryColor,'black','green','red']) 
    .setChartType(Charts.ChartType.COLUMN)
    .setPosition(13,1,0,0)
    .setOption('width', 925)
    .setOption('height', 675)
    .setOption('legend', 'right')
    .setOption('title', engageTitle)
    .setOption('titlePosition', 'out')
    .setOption('titleTextStyle.color','black')
    .setOption('titleTextStyle.fontSize','32')
    .setOption('titleTextStyle.bold','true')
    .setOption('titleTextStyle.italic','false')
    .setOption('hAxis.textStyle.bold','true')
    .setOption('annotations.textStyle.color','black')
    .setOption('annotations.textStyle.fontSize','16')
    .setOption('annotations.textStyle.auraColor','white')
    .setOption('annotations.textStyle.bold','true')
    .setOption('annotations.style','point')
    .setOption('annotations.stem.color','none')
    .setOption('annotations.highContrast', 'true')
    .setOption('annotations.alwaysOutside', 'true')
    .setOption('bar.groupWidth','80%')
    .setOption('isStacked','false')
    .setOption('series.0.dataLabel','value')
    .setOption('series.2.dataLabel','value')
    .setOption('series.1.type','line')
    .setOption('series.1.role','annotation')
    .setOption('series.3.type','line')
    .setOption('series.3.role','annotation')
    .setOption('chartArea.left','100')
    .setOption('chartArea.width','500')
    .setOption('chartArea.top','100')
    .setOption('chartArea.height','500')
    .build();
    engageSheet.insertChart(engageChart);
  }
  
  function copyEngage() {
    //copy the graphs out to the google drive for posting to sites, printing, etc
    var charts = engageSheet.getCharts();
    var chartBlobs = [];
    for (var i in charts) {
      if (!apiEnabled) {
        deleteFile(folder,engageTitle);
      }
      chartBlobs[i]= charts[i].getAs('image/png').setName(engageTitle);
      var fileExist = checkExist(folder,engageTitle);
      if (fileExist) {
        var fileID = getFileId(folder,engageTitle);
        replaceFile(chartBlobs[i], engageTitle, fileID);
      } else {
        folder.createFile(chartBlobs[i]);
      }
    }
  }
}
