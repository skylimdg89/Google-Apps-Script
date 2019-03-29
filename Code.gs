/***
author: dklim
description: 

copy and paste the schdule

***/

//common variables
var spreadsheet = SpreadsheetApp.getActiveSheet(); 
var calendarId = "skylimdg89@gmail.com"; //change calendar id for your gmail id

var lastRow = spreadsheet.getLastRow();
var lastCol = spreadsheet.getLastColumn();
//Logger.log("last ROW = " + lastRow);

var eventCal = CalendarApp.getCalendarById(calendarId);

var numRow = FindNumOfRows(); // call FindNumOfRows function 

var date_column = "A";
var user_column = "B";
var ot_column = "C";
var row_start = "3";

var start = spreadsheet.getRange(date_column + row_start).getValue();
var end = spreadsheet.getRange(date_column+numRow).getValue();

var fromDate = new Date(start);
var toDate = new Date(end);

var calendar = CalendarApp.getCalendarById(calendarId);
var events = calendar.getEvents(fromDate, toDate);

var oncall_string = "(9-18|9-18 on call)";
var off_string = "Off";
var day_string = "Day";
var night_string = "Night";

var start_day_time = "8:00:00 AM"; // =night end time
var end_day_time = "8:00:00 PM"; // =night start time
var end_holiday_night_time = "0:00:00 AM";
var start_oncall_time = "9:00:00 AM";
var end_oncall_time = "6:00:00 PM";
var start_ot_night_time = "10:00:00 PM";
var end_ot_night_time = "6:00:00 AM";
var break_t = "1";

// Edit below variables as needed
var public_holiday_regex = "(Jan 01|Feb 04|Feb 05|Feb 06|Mar 01|May 01|May 06|Jun 06|Aug 15|Sep 12|Sep 13|Oct 03|Oct 09|Dec 25)"//original 2019
var public_holiday_array = ["Jan 01", "Feb 04", "Feb 05", "Feb 06", "Mar 01", "May 06", "Jun 06","Aug 15", "Sep 12", "Sep 13", "Oct 03", "Oct 09", "Dec 25"]; //original 2019
//var public_holiday_regex = "(Mar 01|Mar 12|Mar 15|Mar 16|Mar 17|Mar 20|Mar 21|Mar 31|Apr 01)"//2019 testing
//var public_holiday_array = ["Mar 01", "Mar 12", "Mar 15", "Mar 16", "Mar 17", "Mar 20", "Mar 21","Mar 31", "Apr 01"]; // 2019 testing

//var public_holiday_regex = "(Jan 01|Feb 15|Feb 16|Feb 17|Mar 01|May 01|May 05|May 22|Jun 06|Jun 13|Aug 15|Sep 23|Sep 24|Sep 25|Sep 26|Oct 03|Oct 09|Dec 25)"//2018 testing
//var public_holiday_array = ["Jan 01", "Feb 15", "Feb 16", "Feb 17", "Mar 01", "May 01", "May 05", "May 22", "Jun 06", "Jun 13", "Aug 15", "Sep 23", "Sep 24", 
//                           "Sep 25", "Sep 26", "Oct 03", "Oct 09", "Dec 25"]; //2018 testing


// get shift data from spreadsheet
var shiftdata = spreadsheet.getRange(date_column + row_start + ":" + user_column + numRow).getValues();
var act_working_hours;
var night_work_ot;

/*
synToCalendar()
This function syns shift schedule in spreadsheet to Google Calendar
preconditions: 
- copy and paste date from Sharepoint into A3
- copy and paste schedule from Sharepoint into B3
description:
- Do not run this function more than once. This function will add duplicated schedule to Google calendar
Day: blue(9)
Night: orange (6)
Off: gray (8)
*/
function syncToCalendar(){
  var blue_cal = '9';
  var orange_cal = '6';
  var gray_cal = '8';
  
  for(x=0; x < shiftdata.length; x++){
    var shiftdate = shiftdata[x];
    var date = shiftdate[0];
    var shift = shiftdate[1];
    
    if(shift.match(day_string)){
      eventCal.createAllDayEvent(shift, date).setColor(blue_cal);
    }
    else if(shift==night_string){
      eventCal.createAllDayEvent(shift, date).setColor(orange_cal);
    }
    else{
      eventCal.createAllDayEvent(shift, date).setColor(gray_cal); 
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
  .addItem("syn to calendar", "syncToCalendar")
  .addItem("OT calculator", "otTimeCalc")
  .addSubMenu(ui.createMenu('Delete')
              .addItem('Delete shifts from Calendar', 'clearCalendar'))
  .addToUi();
}

/*
clearCalendar function will remove shift schedule from Google calendar (range: A3 to AnumRow on spreadsheet)

preconditions: 
- enter start date and end date that you want to remove from on spreadsheet A3 to An (n is numRow)
*/
function clearCalendar(){
  var regex = "(Day|Off|Night|9-18|9-18 on call)";
  //var regex = day_string|off_string|night_string|oncall_string;
  for(var i=0; i<events.length;i++){
   var ev = events[i];
    Logger.log(ev.getTitle());
    
    if(ev.getTitle().match(regex)){
      ev.deleteEvent();
    }
  }
}

/*
clearOT function simply clears OT time on spreadsheet 
*/
function clearOT(){
  // range - H2 ~ H(numRow), T2 ~ T(numRow)
  var ot_value;
  var total_ot;
  var copy_helper;
  for(var i = 0; i < numRow; i++){ 
    ot_value = spreadsheet.getRange("H"+(i+3) + ":T"+(i+3)).setValue("");
  }
  total_ot = spreadsheet.getRange("C2").setValue("");
  copy_helper = spreadsheet.getRange("Q"+numRow).setBackground("white");
}

/*
otTimeAdd function calculates ot time by using copied data from Sharepoint
*/
function otTimeAdd(){
  
  var date_string; 
  var tmr_date_string;
    
  var day_start;
  var day_end;
  var night_start;
  var night_end;
  var ot_night_start;
  var ot_night_end;
  var holiday_day_start;
  var holiday_day_end;
  var holiday_night_today_start;
  var holiday_night_today_end;
  var holiday_night_tmr_start;
  var holiday_night_tmr_end;
  
  var oncall_start;
  var oncall_end;
  
  var break_time;
  
  var holiday_night_ct = 0; // debug
  
  var day_ct = 0;
  var night_ct = 0;
  var oncall_ct = 0;
  var holiday_day_ct = 0;
  var holiday_tonight_ct = 0;
  var holiday_tmrnight_ct = 0;
  var night_night_holiday_ct = 0;
  
  var copy_helper;
  
  //Formular used in MS Ecxel to calculate actual working hours and night work ot
  //actual formula requires equal sign(=) but Google spread sheet has the same formula(MINUTE) which is not compatible with MS Excel one
  //so I removed equal sign(=) and return the string result on Googld sheet 
  act_working_hours = spreadsheet.getRange("K" + row_start).setValue('MINUTE(TEXT(MOD(I3-H3,1),"h:m"))/60+HOUR(TEXT(MOD(I3-H3,1),"h:m"))-J3');
  night_work_ot = spreadsheet.getRange("N" + row_start).setValue('MINUTE(TEXT(MOD(M3-L3,1),"h:m"))/60+HOUR(TEXT(MOD(M3-L3,1),"h:m"))');
  
  for(var i=0; i < shiftdata.length; i++){
    var shiftdate = shiftdata[i];
    var date = shiftdate[0];
    var shift = shiftdate[1];
    var tmrdate = shiftdata[i+1];
    date = String(date).split(" ");
    date_string = date[1] + " " + date[2];
    
    tmrdate = String(tmrdate).split(" ");
    tmr_date_string = tmrdate[1] + " " + tmrdate[2];
    
    //act_working_hours = spreadsheet.getRange("K"+(i+3)).setValue('=MINUTE(TEXT(MOD(I'+(i+3)+ '-H'+(i+3) + ',1),"h:m"))/60+HOUR(TEXT(MOD(I'+ (i+3)+ '-H'+(i+3) + ',1),"h:m"))-J'+(i+3));
    
    if(shift.match(day_string)){
      day_ct++;
      
      day_start = spreadsheet.getRange("H"+(i+3)).setValue(start_day_time); //8:00:00 AM
      day_end = spreadsheet.getRange("I"+(i+3)).setValue(end_day_time); //8:00:00 PM
      break_time = spreadsheet.getRange("J"+(i+3)).setValue(break_t);
      
      if(date_string.match(public_holiday_regex)){
        holiday_day_ct++;
        holiday_day_start = spreadsheet.getRange("O"+(i+3)).setValue(start_day_time);//8:00:00 AM
        holiday_day_end = spreadsheet.getRange("P"+(i+3)).setValue(end_day_time);//8:00:00 PM
        break_time = spreadsheet.getRange("Q"+(i+3)).setValue(break_t);
      }
      
    }else if(shift.match(night_string)){
      night_ct++;
      Logger.log("night = " + date_string);
      night_start = spreadsheet.getRange("H"+(i+3)).setValue(end_day_time);//8:00:00 PM
      night_end = spreadsheet.getRange("I"+(i+3)).setValue(start_day_time);//8:00:00 AM
      break_time = spreadsheet.getRange("J"+(i+3)).setValue(break_t);
      
      ot_night_start = spreadsheet.getRange("L"+(i+3)).setValue(start_ot_night_time);//10:00:00 PM
      ot_night_end = spreadsheet.getRange("M"+(i+3)).setValue(end_ot_night_time);//6:00:00 AM
      
      //tmr night is holiday
      if((tmr_date_string.match(public_holiday_regex))){
        holiday_tmrnight_ct++;
        holiday_night_tmr_start = spreadsheet.getRange("O"+(i+3)).setValue(end_holiday_night_time);//0:00:00 AM
        holiday_night_tmr_end = spreadsheet.getRange("P"+(i+3)).setValue(start_day_time);//8:00:00 AM
        break_time = spreadsheet.getRange("Q"+(i+3)).setValue(break_t);
      }
      //
      
      //holiday tonight
      if((date_string.match(public_holiday_regex))){
        holiday_tonight_ct++;
        holiday_night_today_start = spreadsheet.getRange("O"+(i+3)).setValue(end_day_time);//8:00:00 PM
        holiday_night_today_end = spreadsheet.getRange("P"+(i+3)).setValue(end_holiday_night_time);//0:00:00 AM
        
        //tmr night is holiday
        for(var j = 0; j < public_holiday_array.length; j++){
          if(tmr_date_string.match(public_holiday_array[j])){
            Logger.log("tmr is holiday = " + tmrdate); 
            night_night_holiday_ct++;
            holiday_night_tmr_start = spreadsheet.getRange("O"+(i+3)).setValue(end_day_time);//8:00:00 PM
            holiday_night_tmr_end = spreadsheet.getRange("P"+(i+3)).setValue(start_day_time);//8:00:00 AM
          }
        }
      }
    }
    else if(shift.match(oncall_string)){
      oncall_ct++;
      oncall_start = spreadsheet.getRange("H"+(i+3)).setValue(start_oncall_time);//9:00:00 AM
      oncall_end = spreadsheet.getRange("I"+(i+3)).setValue(end_oncall_time);//6:00:00 PM
      break_time = spreadsheet.getRange("J"+(i+3)).setValue(break_t);
    }
      
    else if(shift.match(off_string)){
        Logger.log("off = " + date_string);
    }
    else{
      Logger.log("else = " + shift);
    }
  }
  
  copy_helper = spreadsheet.getRange("Q"+numRow).setBackground("yellow");
  
  var day_total = day_ct - holiday_day_ct;
  var holiday_tonight_total = holiday_tonight_ct - night_night_holiday_ct;
  var holiday_tmrnight_total = holiday_tmrnight_ct - night_night_holiday_ct;
  var night_total = night_ct - holiday_tonight_total - holiday_tmrnight_total - night_night_holiday_ct;
 
  /*
  Logger.log("TOTAL COUNT = " + holiday_day_ct + holiday_tonight_ct + holiday_tmrnight_ct + night_night_holiday_ct + oncall_ct);
  Logger.log("day count(8) = " + day_total);
  Logger.log("night count(0) = " + night_total);
  Logger.log("oncall count(0) = " + oncall_ct);
  Logger.log("holiday day count(2) = " + holiday_day_ct);
  Logger.log("holiday tonight count(3) = " + holiday_tonight_total); 
  Logger.log("holiday tmrnight count(1) = " + holiday_tmrnight_total);
  Logger.log("holiday night_night count(3) = " + night_night_holiday_ct);
  */
  
  var ct_oncall = oncall_ct * 8;
  var ct_day = day_total * 11; //11
  var ct_holiday_day = holiday_day_ct * 22;//22
  var ct_night = night_total * 19; //19
  var ct_holiday_tonight = holiday_tonight_total * 23; //23
  var ct_holiday_tmrnight =  holiday_tmrnight_total * 26; //26
  var ct_night_night_holiday = night_night_holiday_ct * 30; //30
  
  /*
  Logger.log("ct holiday day = " + ct_holiday_day);
  Logger.log("ct holiday tonight" + ct_holiday_tonight);
  Logger.log("ct holiday tmrnight" + ct_holiday_tmrnight);
  Logger.log("ct night night holiday" + ct_night_night_holiday);
  */
  var ot = spreadsheet.getRange(ot_column + "2").setValue(ct_oncall + ct_day + ct_night + ct_holiday_day + ct_holiday_tonight + ct_holiday_tmrnight + ct_night_night_holiday);
  
}


