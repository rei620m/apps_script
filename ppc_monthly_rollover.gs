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
  var ws1 = ss.getSheetByName("Account 1");

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

  var oldMonthCostConv1 = ws1.getRange("I9");
  oldMonthCostConv1.setFormula("=G9/H9");

  var range1 = ws1.getRange(6, 2, 1, data1.length);
    var options = {
    formatOnly: false,
    contentsOnly: true };
   range1.clear(options);

  var ws2 = ss.getSheetByName("Account 2");
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

  var oldMonthCostConv2 = ws2.getRange("I9");
  oldMonthCostConv2.setFormula("=G9/H9");

  var range2 = ws2.getRange(6, 2, 1, data2.length);
    var options = {
    formatOnly: false,
    contentsOnly: true };
   range2.clear(options);

  var ws3 = ss.getSheetByName("Account 3");
  var currentClicks3 = ws3.getRange(6,2).getValue();
  var currentImps3 = ws3.getRange(6,3).getValue();
  var currentCtr3 = ws3.getRange(6,4).getValue();
  var currentCpc3 = ws3.getRange(6,5).getValue();
  var currentConvRate3 = ws3.getRange(6,6).getValue();
  var currentCost3 = ws3.getRange(6,7).getValue();
  var currentConv3 = ws3.getRange(6,8).getValue();
  var data3 = [oldMonth, currentClicks3, currentImps3, currentCtr3, currentCpc3, currentConvRate3, currentCost3, currentConv3];

  ws3.insertRows(9, 1);
  ws3.getRange(9, 1, 1, data3.length).setValues([data3]);
  ws3.getRange(9, 1, 1, data3.length).setBackground("yellow");

  var oldMonthCostConv3 = ws3.getRange("I9");
  oldMonthCostConv3.setFormula("=G9/H9");

  var range3 = ws3.getRange(6, 2, 1, data3.length);
    var options = {
    formatOnly: false,
    contentsOnly: true };
   range3.clear(options);

  var ws4 = ss.getSheetByName("Account 4");
  var currentClicks4 = ws4.getRange(6,2).getValue();
  var currentImps4 = ws4.getRange(6,3).getValue();
  var currentCtr4 = ws4.getRange(6,4).getValue();
  var currentCpc4 = ws4.getRange(6,5).getValue();
  var currentConvRate4 = ws4.getRange(6,6).getValue();
  var currentCost4 = ws4.getRange(6,7).getValue();
  var currentConv4 = ws4.getRange(6,8).getValue();
  var data4 = [oldMonth, currentClicks4, currentImps4, currentCtr4, currentCpc4, currentConvRate4, currentCost4, currentConv4];

  ws4.insertRows(9, 1);
  ws4.getRange(9, 1, 1, data4.length).setValues([data4]);
  ws4.getRange(9, 1, 1, data4.length).setBackground("yellow");

  var oldMonthCostConv4 = ws4.getRange("I9");
  oldMonthCostConv4.setFormula("=G9/H9");

  var range4 = ws4.getRange(6, 2, 1, data4.length);
    var options = {
    formatOnly: false,
    contentsOnly: true };
   range4.clear(options);

  var ws5 = ss.getSheetByName("Account 5");
  var currentClicks5 = ws5.getRange(6,2).getValue();
  var currentImps5 = ws5.getRange(6,3).getValue();
  var currentCtr5 = ws5.getRange(6,4).getValue();
  var currentCpc5 = ws5.getRange(6,5).getValue();
  var currentConvRate5 = ws5.getRange(6,6).getValue();
  var currentCost5 = ws5.getRange(6,7).getValue();
  var currentConv5 = ws5.getRange(6,8).getValue();
  var data5 = [oldMonth, currentClicks5, currentImps5, currentCtr5, currentCpc5, currentConvRate5, currentCost5, currentConv5];

  ws5.insertRows(9, 1);
  ws5.getRange(9, 1, 1, data5.length).setValues([data5]);
  ws5.getRange(9, 1, 1, data5.length).setBackground("yellow");

  var oldMonthCostConv5 = ws5.getRange("I9");
  oldMonthCostConv5.setFormula("=G9/H9");

  var range5 = ws5.getRange(6, 2, 1, data5.length);
    var options = {
    formatOnly: false,
    contentsOnly: true };
   range5.clear(options);

}
