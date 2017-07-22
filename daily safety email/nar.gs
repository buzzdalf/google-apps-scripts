function runNAR() {
  var weekday = isWeekday();
  if (weekday) {
    var startTime = getStartTime();
    var alerts = getNARAlerts(startTime);
    var sewo = getNARSewos(startTime);
    var emails = findEmails();
    sendNAREmail(alerts, sewo, emails);
  }
}

function getNARAlerts(startTime) {
  var alertFolder = '0BxiuqJYDWNpkUWk1c3ZrZnNtVlU';
  var folder = DriveApp.getFolderById(alertFolder);
  var search = 'title contains "Safety Alert" and (modifiedDate > "' + startTime.start + '")';
//  var search = 'title contains "Safety Alert"';

  var alerts = searchFolder(folder,search);
  var text = '';
  if (alerts.length > 0) {
    text = 'NAR had '+alerts.length+' safety alerts in the last 24 hours.  The alerts are listed and attached here:'+'\n';
    for (var i=0;i<alerts.length;i++) {
      text += alerts[i].name+': '+alerts[i].url+'\n';
    }
  } else {
    text = 'NAR did not have any new safety alerts in the last 24 hours. Here is the folder I looked in:'+'\n'+
      'https://drive.google.com/drive/u/0/folders/'+alertFolder+'\n';
  }
  return text;
}

function getNARSewos(startTime) {
  var sewoFolder = '0BxiuqJYDWNpkOFotZmlSWE9Cb3c';
  var folder = DriveApp.getFolderById(sewoFolder);
  var search = 'title contains "SEWO" and (modifiedDate > "' + startTime.start + '")';
  var alerts = searchFolder(folder,search);
//  Logger.log(alerts);
  var text = '';
  if (alerts.length > 0) {
    text = 'NAR had '+alerts.length+' SEWOs in the last 24 hours.  The SEWOs are listed and attached here:'+'\n';
    for (var i=0;i<alerts.length;i++) {
      text += alerts[i].name+': '+alerts[i].url+'\n';
    }
  } else {
    text = 'NAR did not have any SEWOs in the last 24 hours. Here is the folder I looked in:'+'\n'+
      'https://drive.google.com/drive/u/0/folders/'+sewoFolder+'\n';
  }
  return text;
}

function sendNAREmail(alerts, sewo, email) {
  var subject = 'Daily NAR Safety Update';
  var body = 'Here is todays safety update for NAR:'+'\n'+'\n'+
    'Here is the safety alert update'+'\n'+
      alerts+'\n'+
        'Here is the list of recent SEWOs'+'\n'+
          sewo+'\n'+
            'Thank you.'+'\n'+
              '(this is an automated message, if errors please notify your facilitator).';
  
//    Logger.log(email+' '+subject+' '+ body);
  try {
    MailApp.sendEmail(email, subject, body);
  } catch (e) {
    logError(e);
  }
}
