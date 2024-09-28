// Pull last month's gmail receipts into google sheets

function getDeliverooReceipt(senderEmail, subjectPrefix) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("deliveroo");
  var senderEmail = "noreply_at_t_deliveroo_com_abc123.appleid.com"; // Update
  var subjectPrefix = "Your order's in the kitchen";

  var threads = GmailApp.search('from:' + senderEmail + ' subject:' + subjectPrefix + ' after:' + getFirstDayOfPreviousMonth() + ' before:' + getLastDayOfPreviousMonth());
  var messages = GmailApp.getMessagesForThreads(threads);

  // Prepare overwrite
  var numRows = sheet.getLastRow();
  var numColumns = 5;
  if (numRows > 1) {
    var range = sheet.getRange(2, 1, numRows - 1, numColumns);
    range.clearContent();
  }

  // Get the existing formula in column F
  var existingFormulas = sheet.getRange(2, 6, numRows - 1, 1).getFormulas();

  // Process the retrieved emails
  for (var i = 0; i < messages.length; i++) {
    var email = messages[i][0];
    var rawContent = email.getRawContent();
    var dateRegex = /Date: (.*?)(\r?\n)/;
    var dateString = rawContent.match(dateRegex)[1];
    var date = new Date(dateString);
    var formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    var body = email.getPlainBody();
    var total = getDeliverooTotal(body)
    var formattedTotal = formatCurrency(total)
    var category = "外食";
    var memo = "";
    var restaurant = getRestaurantName(body); 
    var row = [formattedDate, category, memo, formattedTotal, restaurant]; 
    var rowIndex = i + 2; // Adjust the row index for writing data
    
    // Overwrite columns A to E
    sheet.getRange(rowIndex, 1, 1, 5).setValues([row]);
    
    // Restore existing formula in column F
    var formula = existingFormulas[i][0];
    sheet.getRange(rowIndex, 6).setFormula(formula);

    var range = sheet.getRange("A2:A20");
    range.setNumberFormat("yyyy/MM/dd");
  }
}

function getDeliverooTotal(body) {
  var regex = /Total\s+\$([\d.]+)/;
  var match = body.match(regex);
  if (match && match[1]) {
    var total = match[1];
    if (!isNaN(parseFloat(total))) {
      return total;
    }
  } else {
    var singleDollarRegex = /\$([\d.]+)/;
    var singleDollarMatch = body.match(singleDollarRegex);
    if (singleDollarMatch && singleDollarMatch[1]) {
      var total = singleDollarMatch[1];
      if (!isNaN(parseFloat(total))) {
        return total;
      }
    }
  }
  return "N/A";
}

function getRestaurantName(body) {
  var keyword = " has your order!";
  var startIndex = body.indexOf(keyword);
  if (startIndex !== -1) {
    var restaurantName = body.substring(0, startIndex);
    restaurantName = restaurantName.trim(); // Remove leading/trailing spaces
    restaurantName = restaurantName.replace(/[\r\n]+/g, ""); // Remove line breaks
    return restaurantName;
  } else {
    return "N/A";
  }
}

function getPaymeEmail(senderEmail) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("payme");
  var senderEmail = "no-reply@secure-app.payme.hsbc.com.hk";

  var threads = GmailApp.search('from:' + senderEmail + ' after:' + getFirstDayOfPreviousMonth() + ' before:' + getLastDayOfPreviousMonth() + ' subject:paid');
  var messages = GmailApp.getMessagesForThreads(threads);

  // Prepare overwrite
  var numRows = sheet.getLastRow();
  var numColumns = 5;
  if (numRows > 1) {
    var range = sheet.getRange(2, 1, numRows - 1, numColumns);
    range.clearContent();
  }

  // Get the existing formula in column F
  var existingFormulas = sheet.getRange(2, 6, numRows - 1, 1).getFormulas();

  // Process the retrieved emails
  for (var i = 0; i < messages.length; i++) {
    var email = messages[i][0];
    var rawContent = email.getRawContent();
    var dateRegex = /Date: (.*?)(\r?\n)/;
    var dateString = rawContent.match(dateRegex)[1];
    var date = new Date(dateString);
      var formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    var subject = email.getSubject();
    var body = email.getPlainBody();
    var total = getPaymeTotal(body)
    var formattedTotal = formatCurrency(total)
    var category = "外食";
    var memo = "";
    var row = [formattedDate, category, memo, formattedTotal, subject]; 
    var rowIndex = i + 2; // Adjust the row index for writing data
    
    // Overwrite columns A to E
    sheet.getRange(rowIndex, 1, 1, 5).setValues([row]);
    
    // Restore existing formula in column F
    var formula = existingFormulas[i][0];
    sheet.getRange(rowIndex, 6).setFormula(formula);

    var range = sheet.getRange("A2:A20");
    range.setNumberFormat("yyyy/MM/dd");
  }
}

function getPaymeTotal(body) {
  var regex = /HKD\s*([^\n<]+)/;
  var match = body.match(regex);
  
  if (match && match[1]) {
    var amount = match[1];
    return amount.trim();
  }

  return "N/A";
}

function formatCurrency(total) {
  var numericTotal = Number(total.replace(/[^0-9.-]+/g, "")); // Remove any non-numeric characters
  var formattedTotal = "$" + numericTotal.toFixed(0); // Format with no decimals and leading dollar sign
  return formattedTotal;
}

function getAppleReceipt(senderEmail, subjectPrefix) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("apple");
  var senderEmail = "no_reply@email.apple.com";
  var subjectPrefix = "Apple からの領収書です";

  var threads = GmailApp.search('from:' + senderEmail + ' subject:' + subjectPrefix + ' after:' + getFirstDayOfPreviousMonth() + ' before:' + getLastDayOfPreviousMonth());
  var messages = GmailApp.getMessagesForThreads(threads);

  // Prepare overwrite
  var numRows = sheet.getLastRow();
  var numColumns = 5;
  if (numRows > 1) {
    var range = sheet.getRange(2, 1, numRows - 1, numColumns);
    range.clearContent();
  }

  // Get the existing formula in column F
  var existingFormulas = sheet.getRange(2, 6, numRows - 1, 1).getFormulas();

  // Process the retrieved emails
  for (var i = 0; i < messages.length; i++) {
    var email = messages[i][0];
    var date = email.getDate();
    var formattedDate = formatDate(date);
    var body = email.getPlainBody();
    var total = getAppleTotal(body);
    var category = "趣味/娯楽";
    var memo = "";
    var hkd = "";
    var row = [formattedDate, category, memo, hkd, total]; 
    var rowIndex = i + 2; // Adjust the row index for writing data
    
    // Overwrite columns A to E
    sheet.getRange(rowIndex, 1, 1, 5).setValues([row]);
    
    // Restore existing formula in column F
    var formula = existingFormulas[i][0];
    sheet.getRange(rowIndex, 6).setFormula(formula);

    var range1 = sheet.getRange("D2:D20");
    range1.setFormula('=IF(A2="", "", INDEX(GOOGLEFINANCE("CURRENCY:JPYHKD", "price", A2), 2, 2) * E2)');
    range1.setNumberFormat("$#,##0");

    var range2 = sheet.getRange("E2:E20"); 
    range2.setNumberFormat("¥#,##0");

    var range3 = sheet.getRange("A2:A20");
    range3.setNumberFormat("yyyy/MM/dd");
  }
}

function getAppleTotal(body) {
  var regex = /¥([\d,]+)/;
  var matches = body.match(regex);

  if (matches && matches.length > 1) {
    var extractedValue = matches[1].replace(/,/g, "").trim();
    var formattedValue = Number(extractedValue);
    return formattedValue;
  }

  return "N/A";
}

// Date helper functions
function formatDate(date) {
  var year = date.getFullYear();
  var month = ("0" + (date.getMonth() + 1)).slice(-2);
  var day = ("0" + date.getDate()).slice(-2);
  return year + "-" + month + "-" + day;
}

function getFirstDayOfPreviousMonth() {
  var today = new Date();
  var firstDayOfThisMonth = new Date(today.getFullYear(), today.getMonth(), 1);
  var firstDayOfPreviousMonth = new Date(firstDayOfThisMonth.getFullYear(), firstDayOfThisMonth.getMonth() - 1, 1);
  return formatDate(firstDayOfPreviousMonth);
}

function getLastDayOfPreviousMonth() {
  var today = new Date();
  var firstDayOfThisMonth = new Date(today.getFullYear(), today.getMonth(), 1);
  var lastDayOfPreviousMonth = new Date(firstDayOfThisMonth.getFullYear(), firstDayOfThisMonth.getMonth(), 0);
  return formatDate(lastDayOfPreviousMonth);
}

function padZero(number) {
  return number < 10 ? '0' + number : number;
}
