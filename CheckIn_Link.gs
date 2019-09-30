function importCheckInReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var agentNamesSheet = ss.getSheetByName('PACE CHART');
  var bmwAgents = agentNamesSheet.getRange('A3:A10').getValues();
  var hondaMiniAgents = agentNamesSheet.getRange('D3:D10').getValues();
  var ignoreNames = ['Jas', 'Mandy'];
  var date = getDate();
  
  if (!date) return;
  
  var metricsCol = parseInt(date.split('/')[1], 10) + 1;
  
  bmwAgents = bmwAgents.map(function (name) {
    return name[0];
  });
  bmwAgents = bmwAgents.filter(function (name) {
    return !!name && ignoreNames.indexOf(name) === -1;
  });
  hondaMiniAgents = hondaMiniAgents.map(function (name) {
    return name[0];
  });
  hondaMiniAgents = hondaMiniAgents.filter(function (name) {
    return !!name && ignoreNames.indexOf(name) === -1;
  });
  
  if (typeof bmwAgents !== 'object') throw 'bmwAgents was not an object!';
  if (typeof hondaMiniAgents !== 'object') throw 'hondaMiniAgents was not an object!';
  
  var metrics = getIndvMetricsCols(metricsCol);
  
  ss.toast('Importing BMW check-in\'s and out\'s now.', 'Importing BMW', 5);
  metrics.values = checkInBMW(bmwAgents, date, metrics);
  ss.toast('Importing Honda/Mini check-in\'s and out\'s now.', 'Importing Honda/Mini', 5);
  metrics.values = checkInHondaMini(hondaMiniAgents, date, metrics);
  
  metrics.range.setValues(metrics.values);
  ss.toast('Import has completed successfully. Have a great day!', 'Import Complete', 5);
}

function getIndvMetricsCols(col) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Individual Metrics');
  var nameValues = sheet.getRange(1, 1, sheet.getLastRow()).getValues();
  var targetRange = sheet.getRange(1, col, sheet.getLastRow());
  var targetValues = targetRange.getValues();
  var formulas = targetRange.getFormulas();
  
  for (var i = 0; i < targetValues.length; i++) {
    if (formulas[i][0]) targetValues[i][0] = formulas[i][0];
  }
  
  nameValues = nameValues.map(function (value) {
    return value[0];
  });
  
  return {
    names: nameValues,
    range: targetRange,
    values: targetValues
  };
}

function checkInBMW(agents, date, metrics) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('CheckInLinkBMW');
  return checkInImport(sheet, agents, date, metrics);
}

function checkInHondaMini(agents, date, metrics) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('CheckInLinkHondaMini');
  return checkInImport(sheet, agents, date, metrics);
}

function checkInImport(sheet, agents, date, metrics) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var agentCell = sheet.getRange('C1');
  var dateCell = sheet.getRange('E1');
  var valuesRange = sheet.getRange('A2:A3');
  var row, values;
  
  if (!date || typeof date !== 'string') throw 'Invalid date! Date needs to be of type "string".';
  if (!agents || typeof agents !== 'object') throw 'Invalid agents! Agents needs to be of type "object".';
  
  dateCell.setValue(date);
  
  for (var i = 0; i < agents.length; i++) {
    agentCell.setValue(agents[i]);
    SpreadsheetApp.flush();
    values = [['Loading...']];
    
    while (values[0][0] === 'Loading...') {
      values = valuesRange.getValues();
      // TODO: Create handler for non-existing sheet
    }
    
    if (values[0][0] === '#REF!') {
      ui.alert('Check In/Out not found',
               'There was an issue locating the sheet for ' + agents[i] + ' on ' + date + '. This agent will be skipped.',
               ui.ButtonSet.OK);
      continue;
    } 
    
    values = {
      leadsSourced: values[0][0],
      itemsSold: values[1][0]
    };
    
    if ((row = metrics.names.indexOf(agents[i])) === -1) throw 'Agent "' + agents[i] + '" could not be found in Individual Metrics sheet!';
    
    metrics.values[row + checkInDriver().offsets.leadsSourced][0] = values.leadsSourced;
    metrics.values[row + checkInDriver().offsets.itemsSold][0] = values.itemsSold;
  }
  
  return metrics.values;
}

function getDate() {
  var ui = SpreadsheetApp.getUi();
  var date, checked;
  var minYear = 2019;
  
  if (ui) {
    while (!checked) {
      date = ui.prompt('Enter Date',
                       'Please enter the date you wish to import in the format (mm/dd/yyyy). Leave blank to use previous working day\'s date.',
                       ui.ButtonSet.OK_CANCEL);
      
      if (date.getSelectedButton() === ui.Button.CANCEL) return;
      
      date = date.getResponseText();
      if (date) {
        if (date.indexOf('/') === -1) {
          ui.alert('Invalid Format', 'The format (' + date + ') is not valid. Please make sure to use "/" to separate the date.', ui.ButtonSet.OK);
          continue;
        }
        
        checked = true;
        date = date.split('/').map(function (item) {
          var int = parseInt(item, 10);
          if (isNaN(int)) checked = false;
          return int;
        });
        
        if (date.length !== 3 || !checked) {
          ui.alert('Error', 'There was an error parsing the date. Please Try again.', ui.ButtonSet.OK);
        } else if (date[0] < 1 || date[0] > 12) {
          checked = false;
          ui.alert('Invalid Month', 'The month entered (' + date[0] + ') is not a valid month.', ui.ButtonSet.OK);
        } else if (date[1] < 1 || date[1] > 31) {
          checked = false;
          ui.alert('Invalid Date', 'The day entered (' + date[1] + ') is not a valid day.', ui.ButtonSet.OK);
        } else if (date[2] < 2019) {
          checked = false;
          ui.alert('Invalid Year', 'The year entered (' + date[2] + ') is not a valid year. Please enter a year of "2019" or later.', ui.ButtonSet.OK);
        }
        
        continue;
      }
      
      checked = true;
    }
  }
  
  if (!date) {
    var milSecPerDay = 1000 * 60 * 60 * 24;
    date = new Date();
    date = new Date(date.getTime() - milSecPerDay);
    
    if (date.getDay() === 0) date = new Date(date.getTime() - milSecPerDay);
    
    date = [date.getMonth() + 1, date.getDate(), date.getFullYear()];
  }
  
  if (typeof date === 'object') date = date.join('/');
  
  if (typeof date !== 'string') throw 'Date is not of type string!';
  
  return date;
}

function checkInDriver() {
  return {
    offsets: {
      itemsSold: 1,
      leadsSourced: 2
    },
  };
}
