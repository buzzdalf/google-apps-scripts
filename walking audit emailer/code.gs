function onEdit(e){
  var ss = e.source.getActiveSpreadsheet();
  var range = e.range;
  if (activeSheet.getName() !== "Structured Table" || e.value !== "notify") return; 

  range.setBackgroundColor('red');
var productname = range.offset(0, -3)
    .getValue();
var productinventory = range.offset(0, -2)
    .getValue();
var message = "Product variant " + productname + " has dropped to " + productinventory;
var subject = "Walking-Working Surface Audit";
var emailAddress = "email@email.com";
MailApp.sendEmail(emailAddress, subject, message);
range.offset(0, 1)
    .setValue("notified");
}
  

function onEdit(e) {
var activeSheet = e.source.getActiveSheet();
var range = e.range;
if (activeSheet.getName() !== "Inventory" || e.value !== "notify") return;
range.setBackgroundColor('red');
var productname = range.offset(0, -3)
    .getValue();
var productinventory = range.offset(0, -2)
    .getValue();
var message = "Product variant " + productname + " has dropped to " + productinventory;
var subject = "Low Stock Notification";
var emailAddress = "email@email.com";
MailApp.sendEmail(emailAddress, subject, message);
range.offset(0, 1)
    .setValue("notified");
}



function myOnEdit (){
  var ss = e.source.getActiveSpreadsheet();
  var tableSource = ss.getSheetByName("Structured Table");
  

// Returns an Array of normalized Strings.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
} 
}