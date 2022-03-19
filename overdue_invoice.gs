// Automated email reminder from alias

function sendEmail() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2; // First row of data to process
  var numRows = 999; // Number of rows to process

  var dataRange = sheet.getRange(startRow, 1, numRows, 999);
  var data = dataRange.getValues();

  // loop through all rows
  for (var i=0; i<numRows; i++) {
    var row = data[i];

  // past due invoice information
  var today = new Date();
  var invoiceDueDate = row[1]; 
  var invoiceStatus = row[8];

  // check for past due invoice
  if (invoiceDueDate < today &&
      invoiceStatus === "Unpaid"
  ) {
      var me = Session.getActiveUser().getEmail();
      var invoiceAlias = GmailApp.getAliases();
      var clientEmail = row[6];

      var invoiceNo = row[2];
      var emailSubject = '[ACTION REQUIRED] OVERDUE INVOICE #' + invoiceNo;

      var firstName = row[4];
      var invoiceUsd = row[3];

      var emailText =
      'Hello ' + firstName + ', ' +
      '\n\n' +
      'This is an automated reminder that your invoice #' + invoiceNo +
      ' for $' + invoiceUsd + ' USD is past due.' +
      '\n' +
      'Please disregard this email if you have already sent the payment, and accept my gratitude.'+
      '\n\n' +
      'If you have any questions, please reach out to me at myemail'+
      '\n\n' +
      'Many thanks,'+
      '\n' +
      'me';
      GmailApp.sendEmail(clientEmail, emailSubject, emailText, {
        from: invoiceAlias[0],
        cc: 'myemail'
      }
      );
    }
  }
}

// run every weekday at 11am
function createTriggers() {
   var days = [ScriptApp.WeekDay.MONDAY, ScriptApp.WeekDay.TUESDAY,
               ScriptApp.WeekDay.WEDNESDAY, ScriptApp.WeekDay.THURSDAY,                                            
               ScriptApp.WeekDay.FRIDAY];
   for (var i=0; i<days.length; i++) {
      ScriptApp.newTrigger("sendEmail")
               .timeBased().onWeekDay(days[i])
               .atHour(11).create();
   }
}
