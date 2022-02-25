// Add OOO dates to personal and team calendar

function addToCalendar() {
  
  function addToPersonal() {
    var spreadsheet = SpreadsheetApp.getActiveSheet();
    var sheet = SpreadsheetApp.getActive().getSheetByName('Sheet1'); 
    var personalCal = CalendarApp.getCalendarById('primary');

    var oooTracker = sheet.getRange(sheet.getLastRow(),1,1,3).getValues();
    var startDate = oooTracker[0][0]; 
    
    personalCal.createAllDayEvent('Rei OOO', startDate);
      var changes = {
      transparency: "transparent"
      };
    };

  function addToTeam() {
    var spreadsheet = SpreadsheetApp.getActiveSheet();
    var sheet = SpreadsheetApp.getActive().getSheetByName('Sheet1'); 
    var teamCal = CalendarApp.getCalendarById('xyz@gmail.com');
    
    var oooTracker = sheet.getRange(sheet.getLastRow(),1,1,3).getValues();
    var startDate = oooTracker[0][0];

    teamCal.createAllDayEvent('Rei OOO', startDate);
      var changes = {
      transparency: "transparent"
      };
  };
}

function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('GAS')
    .addItem('Add to calendar', 'addToCalendar')
    .addToUi();
}
