// Send html email

function createEmail(emails, cc_email) {
  var subject = "subject"
  var emailTemplate = HtmlService.createTemplateFromFile('email_template');
  var body = emailTemplate.evaluate().getContent();
  
  MailApp.sendEmail(
    emails,
    subject,"",{
    cc: cc_email,
    htmlBody: body
    } 
  ); 
}
