var body = DocumentApp.getActiveDocument().getBody();

function onOpen() {
  var ui = DocumentApp.getUi();
  
  ui.createMenu('Alert')
  .addSubMenu(ui.createMenu('New Alert')
              .addItem('New Deviation', 'newDeviation')
              .addItem('New Visual Standard Variance', 'newVariance')
              .addItem('New Unacceptable Parts', 'newUnacceptable'))
  .addSeparator()
  .addSubMenu(ui.createMenu('Extend Alert')
              .addItem('Extend Deviation', 'extendDeviation')
              .addItem('Extend Visual Standard Variance', 'extendVariance')
              .addItem('Extend Unacceptible Parts', 'extendUnacceptable'))
  .addSeparator()
  .addItem('Clear Highlighting', 'clearInputs')
  .addSeparator()
  .addItem('Get Form', 'openForm')
  .addToUi();
}

function newDeviation() {
  getAlert('New Alert','Deviation');
}
function newVariance() {
  getAlert('New Alert','Visual Standard Variance');
}
function newUnacceptable() {
  getAlert('New Alert','Unacceptable Parts');
}
function extendDeviation() {
  getAlert('Extended Alert','Deviation');
}
function extendVariance() {
  getAlert('Extended Alert','Visual Standard Variance');
}
function extendUnacceptable() {
  getAlert('Extended Alert','Unacceptable Parts');
}
function openForm() {
//  var form = FormApp.openByUrl('https://docs.google.com/a/whirlpool.com/forms/d/e/1FAIpQLSeV83uhG8LgMG631VU9D0cAfMiNlkO-AhooTKZHKW0ROAigcQ/viewform');
  showURL("https://docs.google.com/a/whirlpool.com/forms/d/e/1FAIpQLSeV83uhG8LgMG631VU9D0cAfMiNlkO-AhooTKZHKW0ROAigcQ/viewform");
}

function showURL(href){
  var app = UiApp.createApplication().setHeight(50).setWidth(200);
  var link = app.createAnchor('Open Google Form ', href);
  app.add(link);  
  var doc = DocumentApp.getUi();
  doc.showModelessDialog(app,"Open Form");
}

function getAlert(isNew,type) {
  var ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1JsrifxQ8iii03lLHt0y9liabvjCkmdDaZzmj84i4OsQ/edit');
  var sheet = ss.getSheetByName('Dashboard');
  var alert = sheet.getRange("A2").getValue();
  
  clearInputs();
  
  var alertText = body.findText("Document #").getElement().getParent().asParagraph().appendText(' '+alert);
  alertText.setForegroundColor('#FF0000');
  
  var alertType = body.findText(isNew).getElement().getParent().asParagraph();
  alertType.setBackgroundColor('#FFFF00');
  
  var issue = body.findText(type).getElement().getParent().asParagraph();
  issue.setBackgroundColor('#FFFF00');
  
}

function testExisting(alert) {
  var existing = body.findText(alert).getElement().getParent().asParagraph();
  return existing;
}

function clearInputs() {
  body.findText('Deviation').getElement().getParent().asParagraph().setBackgroundColor(null);
  body.findText('Visual Standard Variance').getElement().getParent().asParagraph().setBackgroundColor(null);
  body.findText('Unacceptable Parts').getElement().getParent().asParagraph().setBackgroundColor(null);
  body.findText('New Alert').getElement().getParent().asParagraph().setBackgroundColor(null);
  body.findText('Extended Alert').getElement().getParent().asParagraph().setBackgroundColor(null);
  
  var alertExist = body.findText('Alert-');
  var style = {};
  style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.LEFT;
  style[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
  style[DocumentApp.Attribute.FONT_SIZE] = 10;
  style[DocumentApp.Attribute.BOLD] = true;
  
  if (alertExist) {
    var docNum = alertExist.getElement().getParent().asParagraph();
    var atts = docNum.getAttributes();
    docNum.clear();
    docNum.appendText(' Document #');
    docNum.setAttributes(atts);
    docNum.setAttributes(style);
  }
  
}


