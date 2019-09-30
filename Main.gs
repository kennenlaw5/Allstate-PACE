function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Utilities').addItem('Import Check-in/out', 'importCheckInReport').addToUi();
}