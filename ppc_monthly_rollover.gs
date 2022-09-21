// PPC tracking sheets monthly rollover

function onOpen() {
  // get active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // create menu
  var menu = [{name: "monthlyRollover", functionName: "monthlyRollover"}];

  // add to menu
  ss.addMenu("appsScript", menu);  
}

function monthlyRollover() { 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws1 = ss.getSheetByName("Fitness - GAds search");

  var oldMonth = ws1.getRange(6,1).getValue();

  var currentClicks1 = ws1.getRange(6,2).getValue();
  var currentImps1 = ws1.getRange(6,3).getValue();
  var currentCtr1 = ws1.getRange(6,4).getValue();
  var currentCpc1 = ws1.getRange(6,5).getValue();
  var currentConvRate1 = ws1.getRange(6,6).getValue();
  var currentCost1 = ws1.getRange(6,7).getValue();
  var currentConv1 = ws1.getRange(6,8).getValue();
  var data1 = [oldMonth, currentClicks1, currentImps1, currentCtr1, currentCpc1, currentConvRate1, currentCost1, currentConv1];

  ws1.insertRows(9, 1);
  ws1.getRange(9, 1, 1, data1.length).setValues([data1]);
  ws1.getRange(9, 1, 1, data1.length).setBackground("yellow");

  var range1 = ws1.getRange(6, 2, 1, data1.length);
    var options = {
    formatOnly: false,
    contentsOnly: true };
   range1.clear(options);

  var ws2 = ss.getSheetByName("Fitness - GAds Display");
  var currentClicks2 = ws2.getRange(6,2).getValue();
  var currentImps2 = ws2.getRange(6,3).getValue();
  var currentCtr2 = ws2.getRange(6,4).getValue();
  var currentCpc2 = ws2.getRange(6,5).getValue();
  var currentConvRate2 = ws2.getRange(6,6).getValue();
  var currentCost2 = ws2.getRange(6,7).getValue();
  var currentConv2 = ws2.getRange(6,8).getValue();
  var data2 = [oldMonth, currentClicks2, currentImps2, currentCtr2, currentCpc2, currentConvRate2, currentCost2, currentConv2];

  ws2.insertRows(9, 1);
  ws2.getRange(9, 1, 1, data2.length).setValues([data2]);
  ws2.getRange(9, 1, 1, data2.length).setBackground("yellow");

  var range2 = ws2.getRange(6, 2, 1, data2.length);
    var options = {
    formatOnly: false,
    contentsOnly: true };
   range2.clear(options);

}
