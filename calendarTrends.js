function onOpen() {
  var subMenus = [{name:"Refresh", functionName: "refresh"}];
  SpreadsheetApp.getActiveSpreadsheet().addMenu("Calendar Trends", subMenus);   
}

function refresh() {
   var oneDay = 1000*60*60*24;
   var dayOfWeek = new Array(7);
   var hourOfDay = new Array(24);
   var last30Days = new Array(30);
  
   for (i=0; i<dayOfWeek.length; ++i) dayOfWeek[i] = 0;
   for (i=0; i<hourOfDay.length; ++i) hourOfDay[i] = 0;
   for (i=0; i<last30Days.length; ++i) last30Days[i] = 0;
  
   var currDate = new Date();
   var prevDate = new Date();
   prevDate.setTime(currDate.getTime()-oneDay*30); 
   var userCalendar = CalendarApp.getCalendarById(Session.getUser().getEmail());
   var events = userCalendar.getEvents(prevDate, currDate);

   for (i=0; i<events.length; ++i) {
      var day = new Date(events[i].getStartTime());
      dayOfWeek[day.getDay()] = dayOfWeek[day.getDay()] + 1;
      hourOfDay[day.getHours()] = hourOfDay[day.getHours()] + 1;
   } 

   var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");
   var day = new Date();
   for (i=0; i<last30Days.length; ++i) {
      var events = userCalendar.getEventsForDay(day);
      last30Days[i] = events.length;
      day.setTime(day.getTime()-oneDay);   
      s.setActiveCell("I"+(i+2));
      s.getActiveCell().setValue((day.getMonth()+1)+"/"+day.getDate());
      s.setActiveCell("J"+(i+2));
      s.getActiveCell().setValue(last30Days[i]);      
   }
  
   for (i=0; i<dayOfWeek.length; ++i) {
      s.setActiveCell("D"+(i+2));
      s.getActiveCell().setValue(dayOfWeek[i]);     
   }
  
   for (i=0; i<hourOfDay.length; ++i) {
      s.setActiveCell("G"+(i+2));
      s.getActiveCell().setValue(hourOfDay[i]);     
   }
}
​
