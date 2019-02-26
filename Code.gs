
function scheduleShifts(){
  var spreadsheet = SpreadsheetApp.getActiveSheet();
  var calendarId = "skylimdg89@gmail.com";

  var lastRow = spreadsheet.getLastRow();
  var lastCol = spreadsheet.getLastColumn();
  
  var eventCal = CalendarApp.getCalendarById(calendarId);
  
  var shiftdata = spreadsheet.getRange("A2:B32").getValues();
  
  
  
  //Logger.log("length = ", shiftdata.length);
  //Logger.log("testing...");
  
  for(x=0; x<shiftdata.length; x++){
    var shiftdate = shiftdata[x];
    var date = shiftdate[0];
    var shift = shiftdate[1];
    
    //Logger.log("shiftdata length = ", shiftdata.length);
    //Logger.log(shift);
    //eventCal.createAllDayEvent(shift, date);
    
    
    if(shift.match("Day")){
      eventCal.createAllDayEvent(shift, date).setColor('9');
    }
    else if(shift=="Night"){
      eventCal.createAllDayEvent(shift, date).setColor('6');
    }
    else{
      eventCal.createAllDayEvent(shift, date).setColor('8');
      
    }
    
  }
  
 
}

function onOpen(){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Update to Calendar")
  .addItem("schedule shifts", "scheduleShifts")
  .addSeparatoe()
  .addSubMenu(ui.createMenu('Delete')
              .addItem('Delete shifts from Calendar', 'clearCalendar'))
  .addToUi()
  ;
  
}

function clearCalendar(){
    var spreadsheet = SpreadsheetApp.getActiveSheet();
  var calendarId = "skylimdg89@gmail.com";

  var lastRow = spreadsheet.getLastRow();
  var lastCol = spreadsheet.getLastColumn();
  
  var eventCal = CalendarApp.getCalendarById(calendarId);
  var start = spreadsheet.getRange("A2").getValue();
  var end = spreadsheet.getRange("A32").getValue();
  
  //var fromDate = new Date("3/30/2019");
  //var toDate = new Date("4/1/2019");
  
  var fromDate = new Date(start);
  var toDate = new Date(end);
  var calendarId = "skylimdg89@gmail.com";
  
  var calendar = CalendarApp.getCalendarById(calendarId);
  var events = calendar.getEvents(fromDate, toDate);
  
  var regex = "(Day|Off|Night)";
  
  for(var i=0; i<events.length;i++){
   var ev = events[i];
    Logger.log(ev.getTitle());
    
    if(ev.getTitle().match(regex)){
      ev.deleteEvent();
    }
    
    //ev.deleteEvent();
  }
}
