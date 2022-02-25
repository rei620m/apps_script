// Add OOO dates to personal and team calendar

function addToCalendar() {
  
  var spreadsheet = SpreadsheetApp.getActiveSheet();
  var sheet = SpreadsheetApp.getActive().getSheetByName('Sheet1'); 
  var personalCal = CalendarApp.getCalendarById('primary');
  var teamCal = CalendarApp.getCalendarById('abc@group.calendar.google.com');

  var oooTracker = sheet.getRange(sheet.getLastRow(),1,1,3).getValues();
  var startDate = oooTracker[0][0]; 
    
  personalCal.createAllDayEvent('Rei OOO', startDate);
    var changes = {
    transparency: "transparent"
    };

  teamCal.createAllDayEvent('Rei OOO', startDate);
    var changes = {
    transparency: "transparent"
   };
}

function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('GAS')
    .addItem('Add to calendar', 'addToCalendar')
    .addToUi();
}
