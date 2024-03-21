function selectFirstEmptyRow() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  sheet.setActiveSelection(sheet.getRange("A"+getFirstEmptyRowWholeRow()))
}

function getFirstEmptyRowWholeRow() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getDataRange();
  var values = range.getValues();
  var row = 0;
  for (var row=0; row<values.length; row++) {
    if (!values[row].join("")) break;
  }
  return (row+1);
}

