function analyze(mseFlag) {
  var ss = ssName();
  var inputSheet = inputSheetName();
  var graphSheetName = inputSheet.getRange(32,2).getValue();
  var defaultName = "Graphs";
  var shift = 12;
  var graphSheet = buildSheet(graphSheetName,defaultName); //name of sheet for analysis & graphs
  var sheetName = graphSheet.getName();
  inputSheet.getRange(32,2).setValue(sheetName); //set the Graph label to the result of duplicate test
  var dataSheetName = inputSheet.getRange(31,2).getValue();
  var dataSheet = ss.getSheetByName(dataSheetName);
  var numRows = dataSheet.getLastRow();
  var numCols = dataSheet.getLastColumn();
  var confCount = 0;
  var factorColor = [];
  var groupCount = 0;
  var groupMin;
  var sampleSize = [];
  var subgroupSize = [];
  var levels;    
  var mge = [];
  var average = [];
  var rbar = [];
  var xbar = [];
  var mr = [];
  var mrbar;
  //for all constants the sample size is index + 2.  Sample sizes of 2 - 25 are listed
  var A2Table = [1.88,1.023,0.729,0.577,0.483,0.419,0.373,0.337,0.308,0.285,0.266,0.249,0.235,0.223,0.212,0.203,0.194,0.187,0.18,0.173,0.167,0.162,0.157,0.153];
  var A3Table = [2.659,1.954,1.628,1.427,1.287,1.182,1.099,1.032,0.975,0.927,0.886,0.85,0.817,0.789,0.763,0.739,0.718,0.698,0.68,0.663,0.647,0.633,0.619,0.606];
  var d2Table = [1.128,1.693,2.059,2.326,2.534,2.704,2.847,2.97,3.078,3.173,3.258,3.336,3.407,3.472,3.532,3.588,3.64,3.689,3.735,3.778,3.819,3.858,3.895,3.931];
  var D3Table = [0,0,0,0,0,0.076,0.136,0.184,0.223,0.256,0.283,0.307,0.328,0.347,0.363,0.378,0.391,0.403,0.415,0.425,0.434,0.443,0.451,0.459];
  var D4Table = [3.267,2.574,2.282,2.114,2.004,1.924,1.864,1.816,1.777,1.744,1.717,1.693,1.672,1.653,1.637,1.622,1.608,1.597,1.585,1.575,1.566,1.557,1.548,1.541];
  var B3Table = [0,0,0,0,0.03,0.118,0.185,0.239,0.284,0.321,0.354,0.382,0.406,0.428,0.448,0.466,0.482,0.497,0.51,0.523,0.534,0.545,0.555,0.565];
  var B4Table = [3.267,2.568,2.266,2.089,1.97,1.882,1.815,1.761,1.716,1.679,1.646,1.618,1.594,1.572,1.552,1.534,1.518,1.503,1.49,1.477,1.466,1.455,1.445,1.435];
  var d2 = [];
  var A2 = [];
  var A3 = [];
  var D3 = [];
  var D4 = [];
  var B3 = [];
  var B4 = [];
  var lclR = [];
  var uclR = [];
  var lclA = [];
  var uclA = [];
  var rowOffset = 0;
  var sourceLabel = [];
  var rangeArray = [];
  var rSeries = [];
  var aSeries = [];
  var mrSeries;
  var lclRSeries = [];
  var uclRSeries = [];
  var lclASeries = [];
  var uclASeries = [];
  var uclMRSeries = [];
  var lclXSeries = [];
  var uclXSeries = [];
  var dataSeries = [];
  
  popUp("analyze"); //Launch a dialog to let user know the script is running, so please wait.

  graphSheet.insertColumns(1,(99));
  var i = 1;
  while (dataSheet.getRange(1,i).getValue() != "Data Entry") {
    i++
      if (i > 11) {
        i = numCols;
        break;
      }
  }
  var dataCol = i;
  graphSheet.setActiveCell("H1");
  for (var j = 1;j<=dataCol;j++) {
    var c = 0;
    for (i = 2;i<=numRows;i++) {
      if (j>1 && dataSheet.getRange(i,j).getValue() == dataSheet.getRange(i,(j-1)).getValue()) {
        c++;
      }
    }
    if (c == (numRows - 1)) {
      confCount++;
    } else {
      var useValues = dataSheet.getRange(1,j,numRows,1).getValues();
      factorColor[j] = dataSheet.getRange(1,j).getFontColor();
      graphSheet.getRange(1,(j+shift-confCount),numRows,1).setValues(useValues);
    }
  }
  dataCol = dataCol - confCount;
  var numCols = dataCol - 1;
  var selc = numCols * 3 + 1;    

  calcSubgroups();
  
  function calcSubgroups() {
//    popUp("calcSubgroups");
    var setCell = graphSheet.getRange(1,(selc+shift));
    graphSheet.setActiveSelection(setCell);
    for (var j=numCols;j>=1;j--) {
      sampleSize[j] = 1;
      for (var i=2;i<=numRows;i++) {
        if(graphSheet.getRange(i,(j+shift)).getValue() == graphSheet.getRange((i+1),(j+shift)).getValue()) {
          sampleSize[j]++;
        } else if (sampleSize[j] > 1) {
          groupCount++;
          break;
        }
      }
    }
    //convert sample size to actual subgroup size
    if (groupCount == 0) {
      levels = false;
      groupCount = 1;
    } else {
      levels = true;
    }
    for (j=1;j<=groupCount;j++) {
      subgroupSize[j] = sampleSize[j] / sampleSize[j+1];
    }
    //decide if rolling up tree going forward or not
    if (mseFlag == 1) {
      groupMin = groupCount;
    } else {
      groupMin = 1;
    }
    rollup();
  }

  //rollup the tree
  function rollup() {
//    popUp("rollup");
    var k = 0;
    var selr = numRows + 1;
    var startCol = numCols + shift + 2;
    dataSeries = graphSheet.getRange(2,(startCol-1),(numRows-1),1).getValues();
    //calculate the range and average for each subgroup
    for (var j=groupCount;j>=1;j--) {
      sourceLabel[j] = graphSheet.getRange(1,(j+shift));
      k++;
      mge[j] = [];
      average[j] = [];
      var selc = numCols + k + shift;
      for (var i=0;i<(numRows-1);i+=sampleSize[j]) {
        rangeArray[j] = (k == 1) ? dataSeries.slice(i,(i+sampleSize[j])) : aSeries[j+1].slice(i,(i+sampleSize[j]));
        if (levels) {
          mge[j][i] = calcRange(rangeArray[j]);
        }
        average[j][i] = calcAvg(rangeArray[j]);
      }
      if (levels) {
        rSeries[j] = fillValue(numRows-1,sampleSize[j],mge[j]);
        rbar[j] = calcAvg(rSeries[j]);
      }
      aSeries[j] = fillValue(numRows-1,sampleSize[j],average[j]);
      xbar[j] = calcAvg(aSeries[j]);
      if (mseFlag != 1) {
        if (j == 1) {
          var m = sampleSize[j];
          for (i=m;i<(numRows-1);i+=sampleSize[j]) {
            mr[i] = calcRange(aSeries[j]);
          }
          mrArray = fillValue(numRows-1,sampleSize[j],mr);
          mrbar = calcAvg(mrArray);
        }
      }
      k++;
    }
  }
  
  dotFrequency(); 
  findConstants();
  getLimits();
  
  if (mseFlag == 1) {
    mse(); 
  } else {
    cov();
  }
  
  //build the dot frequency diagram
  function dotFrequency() {
    var selc = numCols + 1 + shift;
    var plotCol = numCols * 3 + shift;
    var plotRow = 1;
    var colCount = subgroupSize[groupCount];
    var dotData = [];
    for (var j=1;j<=colCount;j++) {
      var k = 1;
      var selc2 = plotCol+j;
      dotData[j] = [];
      for (var i=(j-1);i<(numRows-1);i+=colCount) {
        var value = dataSeries[i];
        var tempData = [];
        tempData.push(value);
        dotData[j].push(tempData);
        k++;
      }
      graphSheet.getRange((1+plotRow),selc2,(k-1),1).setValues(dotData[j]);
    }
    graphSheet.getRange(plotRow,(plotCol+1)).setValue("Data formatted for dot frequency");
    var plot = graphSheet.getRange((1+plotRow),(plotCol+1),(k-1),colCount);
    drawDotFreq(plot);
  }
  
  //calculate out the COV
  function cov() {
//    popUp("cov");
    var sigma = [];
    var totalSigma = 0;
    var selr = 5;
    var selc = 7;
    graphSheet.setActiveCell("H1");
    graphSheet.getRange(selr,selc).setValue("Variance Components");
    graphSheet.getRange(selr,(selc+2)).setValue("Percent Contributor");
    for (var j=(groupCount+1);j>=1;j--) {
      var k = 1;
      var n = 0;
      var formatRange = graphSheet.getRange((j+selr),selc);
      formatRange.setValue(graphSheet.getRange(1,(j+shift)).getValue());
      formatRange.setFontWeight("bold");
      if (j == 1) {
        sigma[j] = Math.pow((mrbar/1.128),2);
        for (var m=j;m<=groupCount;m++) {
          k = subgroupSize[m] * k;
          n++;
          sigma[j] = sigma[j]-sigma[j+n]/k;
        }
      } else if (j != (groupCount+1)) {
        sigma[j] = Math.pow((rbar[j-1]/d2[j-1]),2);
        for (var m=j;m<=groupCount;m++) {
          k = subgroupSize[m] * k;
          n++;
          sigma[j] = sigma[j]-sigma[j+n]/k;
        }
      } else {
        sigma[j] = Math.pow((rbar[j-1]/d2[j-1]),2);
      }
      if (sigma[j] < 0) {
        sigma[j] = 0;
      }
      graphSheet.getRange((j+selr),(selc+1)).setValue(sigma[j]);
      totalSigma += sigma[j];
    }
    for (var j=1;j<=(groupCount+1);j++) {
      var formatRange = graphSheet.getRange((j+selr),(selc+2));
      formatRange.setValue(sigma[j]/totalSigma);
      formatRange.setFontSize(12);
      formatRange.setFontWeight("bold");
      formatRange.setNumberFormat("0%");
    }
    formatSheet();
  }
  
  //calculate out the MSE
  function mse() {
//    popUp("mse");
    var testRange;
    var k = 0;
    var j = groupCount;
    var stable = true;
    var between = 0;
    var minRange = mge[j][0];
    var selr = 5;
    var selc = 7;
    graphSheet.setActiveCell("H1");
    graphSheet.getRange((selr-1),selc).setValue("MSE Tests");
    for (var i=0;i<(numRows-1);i+=sampleSize[j]) {
      k++;
      stable = (mge[j][i] > uclR[j]) ? false : stable;
      if (average[j][i] > uclA[j] || average[j][i] < lclA[j]) {
        between++;
      }
      for (var m=0;m<(numRows-1);m+=sampleSize[j]) {
        testRange = mge[j][i] - mge[j][m];
        if (testRange < 0) {
          testRange*=-1;
        }
        if (testRange != 0 && testRange < minRange) {
          minRange = testRange;
        }
      }
    }
    if (minRange == 0) {
      minRange = 1;
    }
    testSPC();
    testDiscrimination();
    testPrecision();
    testBias();
    testSampling();
    formatSheet();
    
    function testSPC() {
//      popUp("testSPC");
      var formatRange = graphSheet.getRange((selr),(selc+1));
      graphSheet.getRange(selr,selc).setValue("Test for SPC");
      if (!stable) {
        formatRange.setValue("Failed. Range Chart Out of Control");
        formatRange.setFontWeight("bold");
        formatRange.setFontColor("red");
      } else {
        formatRange.setValue("Pass. Range Chart In Control");
        formatRange.setFontWeight("bold");
        formatRange.setFontColor("green");
      }
      return;
    }
    
    function testDiscrimination() {
//      popUp("testDisc");
      var discTest;
      var discUnits;
      graphSheet.getRange((selr+2),selc).setValue("Test for Descrimination");
      if (subgroupSize[j] <= 2) {
        discTest = 4;
      }
      if (subgroupSize[j] >= 2 && subgroupSize[j] <= 5) {
        discTest = 5;
      }
      if (subgroupSize[j] == 6) {
        discTest = 6;
      }
      if (subgroupSize[j] > 6) {
        discTest = 5;
      }
      var discMath = (uclR[j]-lclR[j])/minRange;
      if (discMath <= 1) {
        discMath = 0;
      }
      discUnits = discMath + 1;
      discUnits = parseInt(discUnits,10);
      graphSheet.getRange((selr+3),(selc+1)).setValue("Need "+discTest+" units. Have "+discUnits+" units.");
      var formatRange = graphSheet.getRange((selr+2),(selc+1));
      if (discUnits < discTest) {
        formatRange.setValue("Failed. Not enough measurement units.");
        formatRange.setFontWeight("bold");
        formatRange.setFontColor("red");
      } else {
        formatRange.setValue("Pass. Adequate measurement units.");
        formatRange.setFontWeight("bold");
        formatRange.setFontColor("green");
      }
      return;
    }
    
    function testPrecision() {
//      popUp("testPrec");
      graphSheet.getRange((selr+5),selc).setValue("Test for 50% rule");
      graphSheet.getRange((selr+6),selc).setValue("(within/between)");
      var rule = between / k;
      var displayRule = parseInt((rule*100),10);
      graphSheet.getRange((selr+6),(selc+1)).setValue(displayRule+"% points outside control limits");
      var formatRange = graphSheet.getRange((selr+5),(selc+1));
      if (rule<0.5) {
        formatRange.setValue("Failed. Too much within subgroup variation.");
        formatRange.setFontWeight("bold");
        formatRange.setFontColor("red");
      } else {
        formatRange.setValue("Pass. Enough between subgroup variation.");
        formatRange.setFontWeight("bold");
        formatRange.setFontColor("green");
      }
      return;
    }

    function testBias() {
//      popUp("testBias");
      graphSheet.getRange((selr+8),selc).setValue("Test for Bias");
      var biasEffect = false;
      var sigmaBias = rbar[j]/Math.sqrt(sampleSize[1]);
      var uclBias = xbar[1] + 3 * sigmaBias;
      var lclBias = xbar[1] - 3 * sigmaBias;
      for (var i=0;i<(numRows-1);i+=sampleSize[j]) {
        biasEffect = (average[1][i] > uclBias || average[1][i] < lclBias) ? true : biasEffect;
      }
      var formatRange = graphSheet.getRange((selr+8),(selc+1));
      if (biasEffect) {
        formatRange.setValue("Possible "+graphSheet.getRange(1,(shift+1)).getValue() +" bias. Check graphs for difference");
        formatRange.setFontWeight("bold");
        formatRange.setFontColor("red");
      } else {
        formatRange.setValue("No apparent bias effect");
        formatRange.setFontWeight("bold");
        formatRange.setFontColor("green");
      }
      return;
    }
    
    function testSampling() {
      graphSheet.getRange((selr+10),selc).setValue("Test for Sampling");
      var formatRange = graphSheet.getRange((selr+10),(selc+1));
      formatRange.setValue("Did these samples represent the population of interest?");
      formatRange.setFontWeight("bold");
      formatRange.setFontColor("blue");
      return;
    }
  }
  
  // get control limits for charts & MSE
  function getLimits() {
//    popUp("getLimits");
    var k = 0;
    var dataCount = 0;
    for (var j=groupCount;j>=groupMin;j--) {
      var sourceLabel = graphSheet.getRange(1,(j+shift));
      var label = sourceLabel.getValue();
      graphSheet.setActiveSelection(sourceLabel);
      k+=3;
      dataCount++;
      var selcData = numCols + dataCount + shift;
      var selc = numCols * 3 + 5 + k + shift;
      if (levels) {
        graphSheet.getRange(1,(selc+1)).setValue(label+" LCL R");//these are all temp until i find out how to pass array into graph
        graphSheet.getRange(1,(selc+2)).setValue(label+" Range");//these are all temp until i find out how to pass array into graph
        graphSheet.getRange(1,(selc+3)).setValue(label+" UCL R");//these are all temp until i find out how to pass array into graph
        graphSheet.getRange(1,(selc+4)).setValue(label+" LCL Avg");//these are all temp until i find out how to pass array into graph
        graphSheet.getRange(1,(selc+5)).setValue(label+" Average");//these are all temp until i find out how to pass array into graph
        graphSheet.getRange(1,(selc+6)).setValue(label+" UCL Avg");//these are all temp until i find out how to pass array into graph
        lclR[j] = D3[j]*rbar[j];
        uclR[j] = D4[j]*rbar[j];
        lclA[j] = xbar[j] - (A2[j]*rbar[j]);
        uclA[j] = xbar[j] + (A2[j]*rbar[j]);
        lclRSeries[j] = fillValue(numRows-1,sampleSize[j],lclR[j]);
        uclRSeries[j] = fillValue(numRows-1,sampleSize[j],uclR[j]);
        lclASeries[j] = fillValue(numRows-1,sampleSize[j],lclA[j]);
        uclASeries[j] = fillValue(numRows-1,sampleSize[j],uclA[j]);
        var templclR = graphSheet.getRange(2,(selc+1),(numRows-1),1).setValues(lclRSeries[j]);//these are all temp until i find out how to pass array into graph
        var tempR = graphSheet.getRange(2,(selc+2),(numRows-1),1).setValues(rSeries[j]);//these are all temp until i find out how to pass array into graph
        var tempuclR = graphSheet.getRange(2,(selc+3),(numRows-1),1).setValues(uclRSeries[j]);//these are all temp until i find out how to pass array into graph
        var templclA = graphSheet.getRange(2,(selc+4),(numRows-1),1).setValues(lclASeries[j]);//these are all temp until i find out how to pass array into graph
        var tempA = graphSheet.getRange(2,(selc+5),(numRows-1),1).setValues(aSeries[j]);//these are all temp until i find out how to pass array into graph
        var tempuclA = graphSheet.getRange(2,(selc+6),(numRows-1),1).setValues(uclASeries[j]);//these are all temp until i find out how to pass array into graph
        var rLabel = label + ' Range Chart';
        var aLabel = label + ' Average Chart';
        controlChart(rLabel,aLabel,templclR,tempR,tempuclR,templclA,tempA,tempuclA);
      }
      if (j == 1) {
        var uclMR = 3.267 * mrbar;
        var lclX = xbar[1] - 3*(mrbar/1.128);
        var uclX = xbar[1] + 3*(mrbar/1.128);
        graphSheet.getRange(1,(selc+7)).setValue(label+" MR");//these are all temp until i find out how to pass array into graph
        graphSheet.getRange(1,(selc+8)).setValue(label+" UCL MR");//these are all temp until i find out how to pass array into graph
        graphSheet.getRange(1,(selc+9)).setValue(label+" LCL X");//these are all temp until i find out how to pass array into graph
        graphSheet.getRange(1,(selc+10)).setValue(label+" UCL X");//these are all temp until i find out how to pass array into graph
        var uclMRSeries = fillValue(numRows-1,sampleSize[j],uclMR);
        var lclXSeries = fillValue(numRows-1,sampleSize[j],lclX);
        var uclXSeries = fillValue(numRows-1,sampleSize[j],uclX);
        var tempMR = graphSheet.getRange(2,(selc+7),(numRows-1),1).setValues(mrArray); //these are all temp until i find out how to pass array into graph
        var tempuclMR = graphSheet.getRange(2,(selc+8),(numRows-1),1).setValues(uclMRSeries);//these are all temp until i find out how to pass array into graph
        var templclX = graphSheet.getRange(2,(selc+9),(numRows-1),1).setValues(lclXSeries);//these are all temp until i find out how to pass array into graph
        var tempuclX = graphSheet.getRange(2,(selc+10),(numRows-1),1).setValues(uclXSeries);//these are all temp until i find out how to pass array into graph
        var rLabel = label + ' Moving Range Chart';
        var aLabel = label + ' Individuals Chart';
        controlChart(rLabel,aLabel,tempuclMR,tempMR,tempuclMR,templclX,tempA,tempuclX);
      }
      k+=3;
      dataCount++;
    }
  }

  function formatSheet() {
//    popUp("formatSheet");
    graphSheet.setActiveCell("A1");
    closePopup(); //close the dialog, the script is done
  }
  
  function findConstants() {
//    popUp("findConstants");
    for (var j=groupCount;j>=groupMin;j--) {
      d2[j] = d2Table[subgroupSize[j]-2];
      A2[j] = A2Table[subgroupSize[j]-2];
      A3[j] = A3Table[subgroupSize[j]-2];
      D3[j] = D3Table[subgroupSize[j]-2];
      D4[j] = D4Table[subgroupSize[j]-2];
      B3[j] = B3Table[subgroupSize[j]-2];
      B4[j] = B4Table[subgroupSize[j]-2];
    }
  }
  
  //draw control charts
  function controlChart(rLabel,aLabel,lclRSeries,rng,uclRSeries,lclASeries,avg,uclASeries) {
//    popUp("controlChart");
    var Rchart = graphSheet.newChart()
       .addRange(lclRSeries)
       .addRange(rng)
       .addRange(uclRSeries)
       .setChartType(Charts.ChartType.LINE)
       .setOption('title', rLabel)
       .setOption('legend', 'none')
       .setPosition((20+rowOffset), 1, 0, 0)
       .setOption("colors",["red","blue","red"])
       .build();
    var Achart = graphSheet.newChart()
       .addRange(lclASeries)
       .addRange(avg)
       .addRange(uclASeries)
       .setChartType(Charts.ChartType.LINE)
       .setOption('title', aLabel)
       .setOption('legend', 'none')
       .setPosition((20+rowOffset), 6, 0, 0)
       .setOption("colors",["red","blue","red"])
       .build();
    graphSheet.insertChart(Rchart);
    graphSheet.insertChart(Achart);
    rowOffset+=20;
  }
  
  //draw a dot frequency diagram
  function drawDotFreq(plot) {
    var range = plot;
    var dotChart = graphSheet.newChart()
    .addRange(range)
    .setChartType(Charts.ChartType.SCATTER)
    .setOption('title', "Dot Frequency Plot")
    .setOption('useFirstColumnAsDomain', false)
    .setOption('legend', 'none')
    .setPosition(1, 1, 0, 0)
    .build();
    graphSheet.insertChart(dotChart);
  }
}

function fillValue(length,sampleSize,variable) {
//  popUp("fillValue");
  var returnVariable = [];
  for (var i=0;i<length;i+=sampleSize) {
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
  return returnVariable;
}

function calcRange(rangeArray){
//  popUp("calcRange");
  var testArray = rangeArray.filter(function(number) {
    return (!isNaN(parseFloat(number)) && isFinite(number))
  });
  var maxInRange = testArray.sort(function(a,b){return b-a})[0][0];
  var minInRange = testArray.sort(function(a,b){return a-b})[0][0];
  return (maxInRange - minInRange);
}

function calcAvg(rangeArray){
//  popUp("calcAvg");
  var count = 0;
  var total = rangeArray.reduce(function(sum, number) {
    if (!isNaN(parseFloat(number)) && isFinite(number)) {
      count++
    }
    return sum + Number(number)
  }, 0);
  return (total/count);
}


function analyzeMSE() {
  var mseFlag = 1;
  analyze(mseFlag);
}

function analyzeCOV() {
  var mseFlag = 0;
  analyze(mseFlag);
}

function popUp(name) {
  var htmlOutput = HtmlService
  .createHtmlOutputFromFile('Page')
  .setSandboxMode(HtmlService.SandboxMode.IFRAME)
  .setWidth(200)
  .setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Calculating...'+name);
}

function closePopup() {
  var htmlOutput = HtmlService
  .createHtmlOutputFromFile('CloseDialog')
  .setSandboxMode(HtmlService.SandboxMode.IFRAME)
  .setWidth(200)
  .setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Calculating...');
}

