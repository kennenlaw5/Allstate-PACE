function uplistLink(name, range) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var nameCols = [];
  var result = [[0],[0]];
  
  for (var i = 0; i < range[0].length; i++) {
    if (range[0][i].toString().toLowerCase() == 'agent') { nameCols.push(i); }
  }
  
  for (var i = 0; i < nameCols.length; i++) {
    for (var j = 0; j < range.length; j++) {
      if (range[j][nameCols[i]].toLowerCase() == name.toLowerCase()) {
        result[0][0] += parseInt(range[j][nameCols[i]+1], 10);
        result[1][0] += parseInt(range[j][nameCols[i]+2], 10);
      }
    }
  }
  return result;
}

function refreshLink() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var range = ss.getSheetByName('Team Metrics').getRange('P1');
  var value = parseInt(range.getValue(), 10);
  
  if (isNaN(value)) { value = 0; }
  value += 1;
  range.setValue(value);
}
