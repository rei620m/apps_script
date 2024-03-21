function checkEmailsAndSendSummary() {
  var emailAddresses = ['example1@gmail.com', 'example2@gmail.com', 'example3@gmail.com']; // Update - add team's emails here
  var threads = GmailApp.search('from:(example1@gmail.com OR example2@gmail.com OR example3@gmail.com) newer_than:1d'); // Update - sender email and time period

  // Array to store response codes
  var responseCodes = [];

  // Iterate through each email thread
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();

    // Iterate through each message in the thread
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];

      // Get the HTML body of the email
      var body = message.getBody();

      // Extract all URLs from the body
      var urls = extractUrlsFromBody(body);

      // Check the response code of each URL
      for (var k = 0; k < urls.length; k++) {
        var url = urls[k];
        var responseCode = getResponseCode(url);

        // Add non-200 response codes to the array
        if (responseCode !== 200) {
          responseCodes.push(responseCode);
        }
      }
    }
  }

// Send email if there are non-200 response codes
if (responseCodes.length > 0) {
  var summary = 'Summary of broken links:\n\n';

  // Iterate through each email thread
  for (var i = 0; i < threads.length; i++) {
    var thread = threads[i];
    var messages = thread.getMessages();

    // Iterate through each message in the thread
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var sender = message.getFrom();
      var subject = message.getSubject();
      var body = message.getBody();

      // Match URLs using a regular expression
      var urlRegex = /https?:\/\/[^\s<"]+/g;
      var urls = body.match(urlRegex);

      // Check the response code of each URL
      if (urls) {
        for (var k = 0; k < urls.length; k++) {
          var url = urls[k];
          var responseCode = getResponseCode(url);

          var cleanUrl = url.replace(/"$/, "");

          // Add broken links to the summary
          if (responseCode !== 200) {
            summary += 'Sender: ' + sender + '\n';
            summary += 'Subject: ' + subject + '\n';
            summary += 'Link: ' + cleanUrl + '\n';
            summary += 'Response Code: ' + responseCode + '\n\n';
          }
        }
      }
    }
  }

  MailApp.sendEmail({
    to: emailAddresses.join(','),
    subject: 'Trigger emails check - broken links detected',
    body: summary,
    noReply: true
  });
} else if (threads.length === 0) {
    // Send email if no eligible emails are found
    MailApp.sendEmail({
      to: emailAddresses.join(','),
      subject: 'Trigger emails check - no eligible test emails',
      body: 'No eligible test emails were found.'
    });
  } else {
    // Send email if all response codes are 200
    MailApp.sendEmail({
      to: emailAddresses.join(','),
      subject: 'Trigger emails check - ok',
      body: 'All test email links returned a response code of 200.'
    });
  }
}

// Define functions
function extractUrlsFromBody(body) {
  var urls = [];
  var regex = /(https?:\/\/[^\s]+)/g;
  var matches = body.match(regex);

  if (matches) {
    urls = matches;
  }

  return urls;
}

function getResponseCode(url){
  var options = {
     'muteHttpExceptions': true,
     'followRedirects': false
   };
  var statusCode ;
  try {
  statusCode = UrlFetchApp .fetch(url) .getResponseCode() .toString() ;
  }
  
  catch( error ) {
  statusCode = error .toString() .match( / returned code (\d\d\d)\./ )[1] ;
  }

  finally {
  return statusCode ;
  }
}
