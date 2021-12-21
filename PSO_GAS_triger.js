////////////////////////////////////////////////
////////////////Dec 21,2021/////////////////////
//////////////Rentaro Nomura////////////////////
////////////////////////////////////////////////
function myFunction() {
    Logger.log("Initializing...");
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); //obtain current spreadsheet
    var sheet = spreadsheet.getActiveSheet(); //obtain current sheet
   // if(activeSheet.getName() != "団室利用フォーム"){
   //   return;
   // }
  
    var randomColor = "#" + String(Math.floor(Math.random()*16777215).toString(16)); //generate ramdom color code in HEX
  
    //----mode selector----//
    var lastRow = sheet.getLastRow(); //obtain last row
    var mode_target = "C" + String(lastRow)
    var mode_range = sheet.getRange(mode_target);
    var mode = mode_range.getValue();
  
    //----form response loader----//
    if(mode == "予約") { //run if mode is set to reservation mode
      var book_ymd_target = "J" + String(lastRow) //create target cell for booking year-month-date
      var book_start_hm = "K" + String(lastRow) //create target cell for booking start hour-minute
      var book_end_hm = "L" + String(lastRow) //create target cell for booking end hour-minute
      var subscriber_target = "E" + String(lastRow) //create target cell for subscriber name
  
      var range_ymd = sheet.getRange(book_ymd_target);
      var ymd_value = range_ymd.getValue();
  
      var range_start_hm = sheet.getRange(book_start_hm);
      var start_hm_value = range_start_hm.getValue();
  
      var range_end_hm = sheet.getRange(book_end_hm);
      var end_hm_value  = range_end_hm.getValue();
  
      var subscriber_range = sheet.getRange(subscriber_target);
      var subscriber = subscriber_range.getValue();
  
    } else if (mode == "キャンセル"){ //run if mode is set to cancelation mode
      var book_ymd_target = "R" + String(lastRow) //create target cell for booking year-month-date
      var book_start_hm = "S" + String(lastRow) //create target cell for booking start hour-minute
      var book_end_hm = "T" + String(lastRow) //create target cell for booking end hour-minute
      var subscriber_target = "E" + String(lastRow) //create target cell for subscriber name
  
      var range_ymd = sheet.getRange(book_ymd_target);
      var ymd_value = range_ymd.getValue();
  
      var range_start_hm = sheet.getRange(book_start_hm);
      var start_hm_value = range_start_hm.getValue();
  
      var range_end_hm = sheet.getRange(book_end_hm);
      var end_hm_value  = range_end_hm.getValue();
  
      var subscriber_range = sheet.getRange(subscriber_target);
      var subscriber = subscriber_range.getValue();
    }　else {
      Logger.log("UNKNOWN MODE")
      return;
    }
    //----date converter----//
    const year = Utilities.formatDate(ymd_value, 'JST', "yyyy"); //get YEAR
    const month = Utilities.formatDate(ymd_value, 'JST', "MM")-1; //get MONTH
    const date = Utilities.formatDate(ymd_value, 'JST', "dd"); //get DATE
    const start_hour = Utilities.formatDate(start_hm_value, 'JST', "HH");  //get HOUR(start)
    const start_minute = Utilities.formatDate(start_hm_value, 'JST', "mm"); //get MINUTE(start)
    const end_hour = Utilities.formatDate(end_hm_value, 'JST', "HH"); //get HOUR(end)
    const end_minute = Utilities.formatDate(end_hm_value, 'JST', "mm"); //get HOUR(end)
  
    var book_start_date = new Date()
    book_start_date.setFullYear(year)
    book_start_date.setMonth(month)
    book_start_date.setDate(date)
    book_start_date.setHours(start_hour)
    book_start_date.setMinutes(start_minute)
    book_start_date.setSeconds(0)
  
    var book_end_date = new Date()
    book_end_date.setFullYear(year)
    book_end_date.setMonth(month)
    book_end_date.setDate(date)
    book_end_date.setHours(end_hour)
    book_end_date.setMinutes(end_minute)
    book_end_date.setSeconds(0)
    Logger.log("booking starts: " + book_start_date)
    Logger.log("booking ends: " + book_end_date)
  
    var booklist_spreadsheet = SpreadsheetApp.openById('1H7SKnJwSa2VfDyjiAGMwE3RF-VeLJU2bfeoGwcPRlAo'); //search spreadsheet by ID
    var booklist_sheet = booklist_spreadsheet.getSheetByName('1月');
  
  if (mode == "予約"){ //run if mode is set to reservation mode
    //----color filling core----//
    //COLUM
    timetable_start_colum = 1 + 1 + (book_start_date.getDate()-1)*5 //offset 1(stating from 1) and 1(for the colum A) calculate day start colum
    timetable_end_colum = timetable_start_colum + 4 //calculate day end colum
  
    //ROW starting
    timetable_start_row = 3 + (book_start_date.getHours() - 9)*2 //offset 9hours to adjust starting row(3)
    if (book_start_date.getMinutes() >= 30) {
      timetable_start_row=timetable_start_row + 1;
    }
  
    //ROW ending
    timetable_end_row = 3 + (book_end_date.getHours() - 9)*2 
    if (book_end_date.getMinutes() >= 30) {
      timetable_end_row=timetable_end_row + 1;
    }
  
    //fill booked cell
    var timetable_row_diff = 1 + timetable_end_row - timetable_start_row
    var booklist_range = booklist_sheet.getRange(timetable_start_row, timetable_start_colum, timetable_row_diff, 5);
    booklist_range.setBackground(randomColor);
  
    //put the name of a subscriber
    var name_cell_range = booklist_sheet.getRange(timetable_start_row, timetable_start_colum);
    name_cell_range.setValue(subscriber).setFontColor('white').setFontWeight("bold");
    } 
    else if (mode == "キャンセル") { //run if mode is set to cancelation mode
      //----color erasing core----//
      //COLUM
      timetable_start_colum = 1 + 1 + (book_start_date.getDate()-1)*5 //offset 1(stating from 1) and 1(for the colum A) calculate day start colum
      timetable_end_colum = timetable_start_colum + 4 //calculate day end colum
  
      //ROW starting
      timetable_start_row = 3 + (book_start_date.getHours() - 9)*2 //offset 9hours to adjust starting row(3)
      if (book_start_date.getMinutes() >= 30) {
        timetable_start_row=timetable_start_row + 1;
      }
  
      //ROW ending
      timetable_end_row = 3 + (book_end_date.getHours() - 9)*2 
      if (book_end_date.getMinutes() >= 30) {
        timetable_end_row=timetable_end_row + 1;
      }
  
      //fill booked cell
      var timetable_row_diff = 1 + timetable_end_row - timetable_start_row
      var booklist_range = booklist_sheet.getRange(timetable_start_row, timetable_start_colum, timetable_row_diff, 5);
      booklist_range.setBackground(null);
  
      //remove the name of a subscriber
      var name_cell_range = booklist_sheet.getRange(timetable_start_row, timetable_start_colum);
      name_cell_range.deleteCells();
    }
  }