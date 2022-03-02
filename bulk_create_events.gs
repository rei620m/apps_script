// Bulk create calendar events from google sheets
// Reference: https://youtu.be/MOggwSls7xQ

function bulkCreateEvents() {
  
  var spreadsheet = SpreadsheetApp.getActiveSheet();
  var sheet = SpreadsheetApp.getActive().getSheetByName('calendar_events'); 
  var eventCal = CalendarApp.getCalendarById('primary');

  var bulkCreateEvent = spreadsheet.getRange("A2:B100").getValues();

  for (x=0; x<bulkCreateEvent.length; x++) {
    var row = bulkCreateEvent[x];
    var startTime = row[0];
    var endTime = row[1];
  
  eventCal.createEvent('event name', startTime, endTime);
  }
} 
