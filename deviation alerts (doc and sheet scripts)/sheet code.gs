/*
*  This is tied to: https://docs.google.com/spreadsheets/d/1JsrifxQ8iii03lLHt0y9liabvjCkmdDaZzmj84i4OsQ
*/

//function onOpen() {
//  SpreadsheetApp.getUi()
//  .createMenu('Refresh List')
//  .addItem('Catch-Up Deviation Alerts', 'findSheets')
//  .addToUi();
//}

// this function sends the regular summary email
function sendEmail(e) {
  var valueArray = e.values;
  var timestamp = valueArray[0];
  var creator = valueArray[1];
  var number = valueArray[2];
  var isNew = valueArray[3];
  var area = valueArray[4];
  var shift = valueArray[5];
  var contacted = valueArray[6];
  var leader = valueArray[7];
  var impacted = valueArray[8];
  var type = valueArray[9];
  var tda = valueArray[11];
  var url = valueArray[13];
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var responseSheet = ss.getSheetByName('Form Responses 1');
  var setupSheet = ss.getSheetByName('Set up');
  var emailSheet = ss.getSheetByName('email');
  
  if (type == 'Deviation') {
    var part = valueArray[10];
    var staticEmail = setupSheet.getRange("E2").getValue();
  } else {
    var part = valueArray[12];
    var staticEmail = setupSheet.getRange("E1").getValue();
  }

  var emailLookup = emailSheet.getDataRange().getValues();
  
  for (var i=0;i<emailLookup.length;i++) {
    if (emailLookup[i][1] == area) {
      var tempEmail = emailLookup[i][2];
    }
  }

  var replyTo = setupSheet.getRange("E3").getValue();
  var subject = type+' issued.';
  var email = tempEmail+','+staticEmail+','+leader;
  //var email = 'julie_a_rathbun@whirlpool.com';
    
  var body = 'Hello, '+'\n'+'\n'+
    'This is to inform you that a '+type+' has been created for '+area+'.'+'\n'+
    'The document is valid for up to 90 days from '+timestamp+'.'+'\n'+
    'After 90 days this document will no longer be valid and an extension will be required.'+'\n'+'\n'+
      'Creator: '+creator+'\n'+
        'Document Number:'+number+'\n'+
          'New/Extended:'+isNew+'\n'+
            'Area:'+area+'\n'+
              'Shift:'+shift+'\n'+
              'Who was contacted:'+contacted+'\n'+
                'Authorized by:'+leader+'\n'+
                  'Impacted areas:'+impacted+'\n'+
                    'Part/Process:'+part+'\n'+
                      'TDA #:'+tda+'\n'+
                        'Document Link:'+url+'\n'+
                          'Thank you'+'\n';
  
  //Logger.log(email+' '+replyTo+' '+subject+' '+body);
  MailApp.sendEmail(email, replyTo, subject, body);
}

function ninetyDays() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var setupSheet = ss.getSheetByName('Set up');
  var replyTo = setupSheet.getRange("E3").getValue();
  var sheet = ss.getSheetByName('Quality Alert/Deviation');
  var data = sheet.getDataRange().getValues();
  var timeCol = findColumn('Number of Days', data);
  
  for (var i=0;i<data.length;i++) {
    if (data[i][timeCol] > 89 && data[i][timeCol] < 91 ) {
//      Logger.log('equals 90');
      sendReminder(replyTo,data[i]);
    }
  }
}

function sendReminder(replyTo,row) {
  var number = row[3];
  var area = row[5];
  var email = row[2];
  var subject = 'test';
  var body = 'Hello, '+'\n'+'\n'+
    'This is to notify you that the document '+number+' for '+area+' has now reached 90 days of activity and has expired.'+'\n'+
    'Please make arrangements to remove the document from the posting location.'+'\n'+ 
      'If an extension is needed follow the instruction link provided in the document'+'\n'+
      'Thank you'+'\n';
  //Logger.log(email+' '+replyTo+' '+subject+' '+body);
  MailApp.sendEmail(email, replyTo, subject, body);
}

// find columns based on labels and returns the column number
function findColumn(criteria,data) {
  for (i=1;i<=data[0].length;i++) {
    if (data[0][i]) {
      var test = (data[0][i].trim());
    } else {var test = ""};
    if (data[0][i] == criteria || test == criteria) {
      return i;
    }
  }
  return;
}

//deviation
/*
[17-03-17 12:27:19:565 EDT] 0: 3/17/2017 12:27:19
[17-03-17 12:27:19:566 EDT] 1: julie_a_rathbun@whirlpool.com
[17-03-17 12:27:19:566 EDT] 2: Document number
[17-03-17 12:27:19:567 EDT] 3: New
[17-03-17 12:27:19:567 EDT] 4: 241 System
[17-03-17 12:27:19:568 EDT] 5: 1
[17-03-17 12:27:19:569 EDT] 6: VSM
[17-03-17 12:27:19:569 EDT] 7: cher_l_dinan@whirlpool.com
[17-03-17 12:27:19:570 EDT] 8: Line 2
[17-03-17 12:27:19:570 EDT] 9: Deviation
[17-03-17 12:27:19:571 EDT] 10: Part Number or Process Name
[17-03-17 12:27:19:571 EDT] 11: TDA Number
[17-03-17 12:27:19:571 EDT] 12: 
[17-03-17 12:27:19:572 EDT] 13: 
*/

//quality alert
/*
[17-03-17 12:25:50:578 EDT] 0: 3/17/2017 12:25:50
[17-03-17 12:25:50:578 EDT] 1: julie_a_rathbun@whirlpool.com
[17-03-17 12:25:50:579 EDT] 2: Document number
[17-03-17 12:25:50:579 EDT] 3: New
[17-03-17 12:25:50:579 EDT] 4: Alpha Drum Line
[17-03-17 12:25:50:580 EDT] 5: 1
[17-03-17 12:25:50:580 EDT] 6: Area Leader
[17-03-17 12:25:50:580 EDT] 7: charles_r_belcher@whirlpool.com
[17-03-17 12:25:50:581 EDT] 8: Line 3
[17-03-17 12:25:50:581 EDT] 9: Quality Alert
[17-03-17 12:25:50:581 EDT] 10: 
[17-03-17 12:25:50:582 EDT] 11: 
[17-03-17 12:25:50:582 EDT] 12: Part number or Process name
[17-03-17 12:25:50:582 EDT] 13: 
*/