// Share a random joke from a list in google sheets, via email

function randomJoke() {
var t = SpreadsheetApp.getActive().getSheetByName('jokes').getRange('A:A').getValues().filter(String)
return t[getRandomInt(0, t.length-1)]
}

function getRandomInt(min, max) {
    return Math.floor(Math.random() * (max - min + 1)) + min;
}

function sendMail () {
  var emailText = randomJoke();
  MailApp.sendEmail('youremail', 'subject', emailText);
}
