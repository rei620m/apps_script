// Send html email

function createEmail(recipients, cc_email) {
  var subject = "subject"
  var emailTemplate = HtmlService.createTemplateFromFile('email_template');
  var body = emailTemplate.evaluate().getContent();
  
  MailApp.sendEmail({
    to: recipients,
    cc: cc_email,
    subject: subject,
    htmlBody: body
    } 
  ); 
}
