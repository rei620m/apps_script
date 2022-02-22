// Receive email alert 3 days before subscription auto renew

function emailAlert() {
  
  // today's date information
  var today = new Date();
  var todayMonth = today.getMonth() + 1;
  var todayDay = today.getDate();
  var todayYear = today.getFullYear();

  // three days
  var threeDaysFromToday = new Date();
  threeDaysFromToday.setDate(threeDaysFromToday.getDate() + 3);
  var threeDaysMonth = threeDaysFromToday.getMonth() + 1;
  var threeDaysDay = threeDaysFromToday.getDate();
  var threeDaysYear = threeDaysFromToday.getFullYear();
  
  // getting data from spreadsheet
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2; // First row of data to process
  var numRows = 999; // Number of rows to process

  var dataRange = sheet.getRange(startRow, 1, numRows, 999);
  var data = dataRange.getValues();

  //looping through all of the rows
  for (var i=0; i<numRows; i++) {
    var row = data[i];

  //next payment date information
    var nextPaymentDateMonth = new Date(row[7]).getMonth() + 1;
    var nextPaymentDateDay = new Date(row[7]).getDate();
    var nextPaymentDateYear = new Date(row[7]).getFullYear();

  //check for 3 days
    if (
      nextPaymentDateMonth === threeDaysMonth &&
      nextPaymentDateDay === threeDaysDay &&
      nextPaymentDateYear === threeDaysYear
    ) {
      var subject = 'Reminder - subscription auto renew';
      var message =
      'Reminder - Your ' +
      row[4] +
      ' subscription for '+
      row[1] + 
      ' '+
      row[0] +
      ' ('+
      row[3] +
      ' JPY), will auto renew in 3 days.';
      MailApp.sendEmail('youremail', subject, message);
    }
  } 
}
