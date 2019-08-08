function newMonth() {
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var sheet       = ss.getSheetByName('Individual Metrics');
  var range       = sheet.getRange(2, 1, sheet.getLastRow() - 1, 32);
  var values      = range.getValues()
  var formulas    = range.getFormulas();
  var rowsToClear = ['items sold', 'total leads sourced', 'mtd hh sold', 'calls made'];
  
  for (var i in values) {
    if (rowsToClear.indexOf(values[i][0].toString().toLowerCase()) !== -1) {
      for (var j = 1; j < values[i].length; j++) values[i][j] = '';
    }
  }
  
  for (i in values) {
    for (j in values[i]) {
      if (formulas[i][j] !== '' && formulas[i][j] !== null && formulas[i][j] !== undefined) {
        values[i][j] = formulas[i][j]
      }
    }
  }
  
  range.setValues(values);
  
  range     = sheet.getRange(1, 2);
  var date  = range.getValue();
  var year  = date.getYear();
  var month = date.getMonth() + 1;
  
  if (month > 11) {
    month = 0;
    year ++;
  }
  
  date = new Date(year, month);
  range.setValue(date);
}

function endOfMonth() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('UplistLink');
  var range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
  range.setValues(range.getValues());
  sheet.hideSheet();
  SpreadsheetApp.flush();
}
