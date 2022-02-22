// Delete all day events in a given time period

function delete_events()
{
  var fromDate = new Date(2022,0,19,0,0,0); // Inclusive 
  var toDate = new Date(2022,0,20,0,0,0); // Exclusive
  var calendarID = 'primary'; 
  var calendar = CalendarApp.getCalendarById(calendarID);
  
var events = calendar.getEvents(fromDate, toDate);
    for(var i=0; i<events.length;i++){
      var ev = events[i];
      Logger.log(ev.getTitle());
      ev.deleteEvent();
  }
}
