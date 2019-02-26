/***
common variables
***/
var spreadsheet = SpreadsheetApp.getActiveSheet();
var calendarId = "skylimdg89@gmail.com";

var lastRow = spreadsheet.getLastRow();
var lastCol = spreadsheet.getLastColumn();

var eventCal = CalendarApp.getCalendarById(calendarId);

var numRow = FindNumOfRows();

var lastRow = spreadsheet.getLastRow();
var lastCol = spreadsheet.getLastColumn();

var start = spreadsheet.getRange("A2").getValue();
var end = spreadsheet.getRange("A"+numRow).getValue();

var fromDate = new Date(start);
var toDate = new Date(end);

var calendar = CalendarApp.getCalendarById(calendarId);
var events = calendar.getEvents(fromDate, toDate);

var public_holidays = [
  "1/1/2019", 
  "2/4/2019",
  "2/5/2019",
  "2/6/2019",
  "3/1/2019",
  "5/1/2019",
  "5/6/2019",
  "6/6/2019",
  "8/15/2019",
  "9/12/2019",
  "9/13/2019",
  "10/3/2019",
  "10/9/2019",
  "12/25/2019"
];

function scheduleShifts(){

  var shiftdata = spreadsheet.getRange("A2:B"+numRow).getValues();
  
  for(x=0; x<shiftdata.length; x++){
    var shiftdate = shiftdata[x];
    var date = shiftdate[0];
    var shift = shiftdate[1];
    
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


function FindNumOfRows() {
  range = SpreadsheetApp.getActiveSheet().getLastRow();
  return range;
}


function onOpen(){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Update to Calendar")
  .addItem("schedule shifts", "scheduleShifts")
  .addItem("OT calculator", "otTimeCals")
  .addSeparatoe()
  .addSubMenu(ui.createMenu('Delete')
              .addItem('Delete shifts from Calendar', 'clearCalendar'))
  .addToUi();
}

function clearCalendar(){

  var regex = "(Day|Off|Night)";
  
  for(var i=0; i<events.length;i++){
   var ev = events[i];
    Logger.log(ev.getTitle());
    
    if(ev.getTitle().match(regex)){
      ev.deleteEvent();
    }
  }
}

function otTimeCalc(){
  
  Logger.log("OT Calculating...");  
  Logger.log("public holidays " + public_holidays);
  var day_ct = 0;
  var night_ct = 0;
  
  for(var i=0; i<events.length;i++){
   var ev = events[i];
    Logger.log(ev.getTitle());
    
    if(ev.getTitle().match("Day")){
      day_ct++;
    }else if(ev.getTitle().match("Night")){
      night_ct++;
    } 
  }
  
  var ot_day = day_ct * 11;
  var ot_night = night_ct * 19;
  var ot = spreadsheet.getRange("C2").setValue(ot_day + ot_night);

}
