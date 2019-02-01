function newMonth() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
}

function endOfMonth() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('UplistLink');
  var range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
  range.setValues(range.getValues());
  sheet.hideSheet();
  SpreadsheetApp.flush();
}
