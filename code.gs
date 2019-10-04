chools_timeCounts = [];
var schools_names = [];

// ** Interpreters
// Number of meetings for all time slots in week for each interp
var interps_timeCounts = [];
var interps_names = [];

/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Import school requests...', functionName: 'importSchool'},
    {name: 'Import interpreter availability...', functionName: 'importInterp'}
  ];
  spreadsheet.addMenu('IHC', menuItems);
}

// Initialize global variables
function init() {
  
  // 10 arrays for 10 schools
  // Each array tracks the number of weekly meetings per time slot
  // 12 slots per day, 5 days per week = 60 time slots per school
  for(var i = 0; i < 10; i++) {
    var arr = new Array(60);
    for (var k = 0; k < 60; k++) arr[k] = 0;
    schools_timeCounts.push(arr);
  }
  
}

/**
 * Update schools_timeCounts with new meeting at time slot for school
 * 
 * school - number for school_timeCounts index
 * time - time of meeting (0:00-0:00 PM)
 * day - num 0-4 to indicate Mon-Fri
 */
function add_schoolsTimeCounts(school, time, day) {
  var times = [ "12:00 - 12:30 PM", "12:30 - 1:00 PM", "1:00 - 1:30 PM", "1:30 - 2:00 PM", 
               "2:00 - 2:30 PM", "2:30 - 3:00 PM", "3:00 - 3:30 PM", "3:30 - 4:00 PM",
               "4:00 - 4:30 PM", "4:30 - 5:00 PM", "5:00 - 5:30 PM", "5:30 - 6:00 PM" ];
  
  s = schools_timeCounts[school];
  time_idx = times.indexOf(time);
  
  // set time_idx to time in day
  time_idx = day*12 + time_idx;
  
  // update
  s[time_idx] = s[time_idx] + 1;
  schools_timeCounts[school] = s;  
}

/**
 * A function that imports and organizes meeting requests per teacher
 * from a maximum of 10 schools.
 */
function importSchool() {
  
  // Initialize global variables
  init();
  
  // Clear any previous meetings in Scheduler
  clearMeetings();
  
  // Scheduler
  var sched = SpreadsheetApp.getActive();
  
  // School spreadsheets imported by ID.
  // Spreadsheet ID is between /d/ and /edit in Sheet URL
  var s1 = SpreadsheetApp.openById("1KDTmKj9K7vmzkNT4K27FZ0eHSSWBrT5rmr85SoPlfQc");
  var s2 = SpreadsheetApp.openById("1ezgpnmgq21vzLTzqFVDpYkcCGFQTklwD7qpGA0GOB9s");
  var s3 = SpreadsheetApp.openById("1lH08Y7-orYnIJo2jYXCc2NCFDCIPa-FKGr6y-QDEtSs");
  var s4 = SpreadsheetApp.openById("1zuImk0BkEWiCt4CEItomupLF-fcJll0z7YwhlNKhfUA");
  var s5 = SpreadsheetApp.openById("1W0l_7UOPZs9bW0LF4ENK-XLlYxyufzN68Znc30gNI5Q");
  var s6 = SpreadsheetApp.openById("1k6zSq9SZJ0qzvgTzueDzJR2BB9P6c0PnmUFQK8F29JM");
  var s7 = SpreadsheetApp.openById("1LjPIV-YIqunXbklzrFOnQT3VlEc1hom5IpESXxhdGIg");
  var s8 = SpreadsheetApp.openById("109Gmcnj1KMIZBS90AZCRiIObs3cL57W_fOx6gWZGXdA");
  var s9 = SpreadsheetApp.openById("10Hpz65vFN3J_LGJ22FdHZ5NUCrOzbGkFkgGrrrzncfg");
  var s10 = SpreadsheetApp.openById("1DKBDln7cZZtdIy82sd9MKbsIypfrEnnccHKFYjx-GjQ");
  
  var schools = [ s1, s2, s3, s4, s5, s6, s7, s8, s9, s10 ];
  
  // 1 teach_rm array per school
  var teachers_all = [[],[],[],[],[],[],[],[],[],[]];
  
  // Use index of time slot to determine offset from teach_cell
  var times = [ "12:00 - 12:30 PM", "12:30 - 1:00 PM", "1:00 - 1:30 PM", "1:30 - 2:00 PM", 
               "2:00 - 2:30 PM", "2:30 - 3:00 PM", "3:00 - 3:30 PM", "3:30 - 4:00 PM",
               "4:00 - 4:30 PM", "4:30 - 5:00 PM", "5:00 - 5:30 PM", "5:30 - 6:00 PM" ];
  
  // Time slot columns for each day in School Request
  var day_cols = [ "A", "E", "I", "M", "Q" ];

  var s;
  var s_name;
  var sheet;
  var num;
  var s_num;
  var teach;
  var rm;
  var student;
  var time;
  var teach_rm;
  var teach_idx;
  var time_col;
  var time_cell;
  var teach_cell;
  var meeting_row;
  var meeting_col;
  
  // For School Request 1-10
  for(var i = 0; i < 10; i++) {
    
    // get SchoolN sheet from Scheduler
    num = i + 1;
    s_num = 'School'.concat(num.toString());
    sheet = sched.getSheetByName(s_num);
    
    // School Request
    s = schools[i];
    teach_rm = teachers_all[i];
    
    // enter school name on Scheduler
    s_name = s.getRange('A1').getValue();
    sheet.getRange('A1').setValue(s_name);
    
    // collect school name for global var
    schools_names.push(s_name);
    
    // For each day in School Request
    for(var j = 0; j < 5; j++) {
      
      time_col = day_cols[j];
      
      // For each time slot in day
      for(var r = 5; r <= 147; r += 2) {
        
        time_cell = time_col.concat(r.toString());
        
        // Meeting info from School Request
        time = s.getRange(time_cell).getValue();
        teach = s.getRange(time_cell).offset(0,1).getValue();
        rm = s.getRange(time_cell).offset(0,2).getValue();
        student = s.getRange(time_cell).offset(0,3).getValue();
        
        // Skip meeting if teacher, time, or student blank
        if (teach == "" || time == "" || student == "") {
         continue; 
        }
        
        // add meeting to schools_timeCounts
        add_schoolsTimeCounts(i, time, j);
        
        
        // Paste meeting to Scheduler
        
        // Index of (teacher, room) in teach_rm determines
        // teacher table spot in Scheduler
        
        teach_idx = teach_rm.indexOf((teach,rm));
        
        // (teacher, room) not in teach_rm
        if (teach_idx == -1) {
          teach_idx = teach_rm.push((teach,rm)) - 1;
        }
        
        // Set teach_cell for corresponding teacher in Scheduler
        teach_idx = 39*teach_idx + 5;
        teach_cell = 'B'.concat(teach_idx.toString());
        
        // Paste Teacher and Room to Scheduler
        sheet.getRange(teach_cell).setValue(teach);
        sheet.getRange(teach_cell).offset(1,0).setValue(rm);
        
        // row and col offsets from teach_cell for meeting paste
        meeting_row = 3*times.indexOf(time) + 2;
        meeting_col = j + 1;
        
        // Paste student to Scheduler at meeting time
        sheet.getRange(teach_cell).offset(meeting_row,meeting_col).setValue(student);
        
      }
    
    }
  }
  
}

/**
 * A function that goes through all School sheets in Scheduler and clears
 * every teacher's name, room num and all appointments (student and interpreter).
 */
function clearMeetings() {
  
  // Scheduler
  var sched = SpreadsheetApp.getActive();
  
  var num;
  var s_num;
  var sheet;
  var table;
  var start_row;
  var stop_row;
  var start;
  var stop;
  var rng;
  var teach_row;
  var teach_cell;
  
  // For School Request 1-10
  for(i = 0; i < 10; i++) {
    
    // get SchoolN sheet from Scheduler
    num = i + 1;
    s_num = 'School'.concat(num.toString());
    sheet = sched.getSheetByName(s_num);
    
    // First teacher table starts row 7
    table = 7;
    
    // Delete all teacher names, room nums and 
    // meetings (student and interp) table by table
    while (table <= 1333) {
      
      // teacher name cell in table
      teach_row = table - 2;
      teach_cell = 'B'.concat(teach_row.toString());
      
      // delete teacher name and room num
      sheet.getRange(teach_cell).clearContent();
      sheet.getRange(teach_cell).offset(1,0).clearContent();
      
      // start and stop rows of noon meeting range
      start_row = table;
      stop_row = table + 1;
      
      // Delete all 12 meetings in current table
      for(j = 0; j < 12; j++) {
        start = 'C'.concat(start_row.toString());
        stop = 'G'.concat(stop_row.toString());
        rng = start.concat(':').concat(stop);
        
        sheet.getRange(rng).clearContent();
        
        start_row += 3;
        stop_row += 3;
      }
    
      // next teacher table
      table += 39;
    }
    
  }
  
}


/**
 * A function that transfers interpreter availabilities from Form Requests
 * sheet to Mon-Fri sheets in workbook. Deletes prior availabilities on Mon-Fri
 * sheets and interpreter data on 'Interpreters' sheet before transferring. The
 * interpreter's most recent submission on 'Form Requests' is transferred to Mon-Fri
 * availability sheets.
 *
 * *Assume Form Responses is first sheet in workbook*
 * *Assume interps spell their name correctly with each submission*
 * *Assume name of 'Interpreters' sheet remains unchanged*
 */
function importInterp() {

  // Scheduler
  var sched = SpreadsheetApp.getActive();
  
  // Clear interpreter data from sheets Mon-Fri
  clearInterps();
  
  // ** Assume Form Responses is the first sheet in workbook. **
  var resp = sched.getSheets()[0];
  
  // Range of all cells with content in column A
  var form_rng = resp.getRange("A1").getDataRegion(SpreadsheetApp.Dimension.ROWS).getA1Notation();
  
  // Last row with interpreter response
  // Equals 1 if no data entries under title in column A
  var name_endRow = form_rng.substring(form_rng.lastIndexOf(':')+2)
  
  var name;
  var names = [];
  
  // Collect max row or most recent submission of each interp
  // [(name, max row)]
  var names_row = [];
  
  // Collect all non-empty unique names and their max row
  // Skip col header at B1
  for(var i = 2; i <= name_endRow; i++) {
    name = resp.getRange('B'.concat(i.toString())).getValue();
    
    // skip blanks
    if (name == "") {
     continue; 
    }
    
    // add unique name at row
    if (names.lastIndexOf(name) == -1) {
      names.push(name);
      names_row.push([name,i]); 
    }
    // update existing name with new row
    else {
      names_row[names.lastIndexOf(name)] = [name,i];
    }
  }
  
  // Transfer each interpreters' availability to Mon-Fri sheets
  
  var max_row;
  var phone;
  var email;
  var perm;
  var car;
  var day_avails;
  
  var avail_cols = [ 'E', 'F', 'G', 'H', 'I' ];
  var day_sheets = [ 'Mon', 'Tues', 'Wed', 'Thur', 'Fri' ];
  
  var interps = sched.getSheetByName('Interpreters');
  var interp_row;
  
  // For each unique interpreter and their max row
  for(var k = 0; k < names_row.length; k++) {
    
    name = names_row[k][0];
    max_row = names_row[k][1];
    phone = resp.getRange('C'.concat(max_row.toString())).getValue();
    email = resp.getRange('K'.concat(max_row.toString())).getValue();
    perm = resp.getRange('D'.concat(max_row.toString())).getValue();
    car = resp.getRange('J'.concat(max_row.toString())).getValue();
    
    // Add asterisk to name if interpreter has transportation
    if (car == "Yes") {
      name += '*';
    }
    
    // Collect interpreter's name for global var
    interps_names.push(name);
    
    // Add time slot array for new interp in global var
    var arr = new Array(60);
    for (var n = 0; n < 60; n++) arr[n] = 0;
    interps_timeCounts.push(arr);
    
    // Transfer interpreter info to 'Interpreters' Sheet
    interp_row = k + 3;
    interps.getRange('A'.concat(interp_row.toString())).setValue(name);
    interps.getRange('B'.concat(interp_row.toString())).setValue(phone);
    interps.getRange('C'.concat(interp_row.toString())).setValue(email);
    interps.getRange('D'.concat(interp_row.toString())).setValue(perm);
    
    // For Mon-Fri availabilities of interpreter
    for (var j = 0; j < 5; j++) {
      day_avails = resp.getRange(avail_cols[j].concat(max_row.toString())).getValue();
      day_avails = day_avails.split(", ");
      
      // Transfer day availability from 'Form Responses' to Mon-Fri sheet
      transAvail(name, day_avails, day_sheets[j]);
    }

  }
  
}

/**
 * A function that transfers single day availability from 'Form Response' to
 * Mon-Fri sheet.
 *
 * avails - Array of times available for day
 * day_sheet - Mon-Fri sheet name for transfer
 */
function transAvail(name, avails, day_sheet) {
  
  // Scheduler and Day Sheet
  var sched = SpreadsheetApp.getActive();
  var sheet = sched.getSheetByName(day_sheet);
  
  // Use index of time slot to determine which col to paste in Day
  var times = [ "12:00-12:30 PM", "12:30-1:00 PM", "1:00-1:30 PM", "1:30-2:00 PM", 
               "2:00-2:30 PM", "2:30-3:00 PM", "3:00-3:30 PM", "3:30-4:00 PM",
               "4:00-4:30 PM", "4:30-5:00 PM", "5:00-5:30 PM", "5:30-6:00 PM" ];
  
  // Data columns in Mon-Fri sheets
  var cols = [ "A", "M", "Y", "AK", "AW", "BI", "BU", "CG", "CS", "DE", "DQ", "EC" ];
  
  var time;
  var paste_col;
  var paste_row;
  var col_rng;
  
  // Paste all availabilities at end of corresponding column in Day sheet
  for(var i = 0; i < avails.length; i++) {
   
    time = avails[i];
    
    // Edge case - avail time not found in times
    if (times.indexOf(time) == -1) {
      continue;
    }
    else {
      paste_col = cols[times.indexOf(time)];
    }
    
    // Range of paste column
    col_rng = sheet.getRange(paste_col.concat('1')).getDataRegion(SpreadsheetApp.Dimension.ROWS).getA1Notation();
    
    // Paste one row past data
    paste_row = sheet.getRange(col_rng).getLastRow() + 1;
    
    // Paste
    sheet.getRange(paste_col.concat(paste_row)).setValue(name);
  }
}


/**
 * A function that clears all imported interpreter availabilities on sheets Mon-Fri
 * and clears their lookup values in 'Interpreters' sheet.
 *
 * Assumes name of 'Interpreters' sheet remains unchanged.
 */
function clearInterps() {

  // Scheduler
  var sched = SpreadsheetApp.getActive();
  
  // Days Monday - Friday
  var week_days = [ "Mon", "Tues", "Wed", "Thur", "Fri" ];
  
  // Data columns in Mon-Fri sheets
  var cols = [ "A", "M", "Y", "AK", "AW", "BI", "BU", "CG", "CS", "DE", "DQ", "EC" ];
  
  
  // Clear 'Interpreters' sheet
  
  // 'Interpreters' Sheet
  var interps = sched.getSheetByName('Interpreters');
  
  // Get full square data range on 'Interpreters' sheet
  var interp_rng = interps.getRange("A1").getDataRegion().getA1Notation();
  
  // Last row with data on 'Interpreters' sheet
  var interp_endRow = interp_rng.substring(interp_rng.lastIndexOf(':')+2);
  
  // Delete data on 'Interpreters' sheet below row 2 (keep dummy vars on row 2)
  if (interp_endRow > 2) {
    interps.getRange('A3:D'.concat(interp_endRow.toString())).clearContent();
  }
  
  
  // Clear Mon-Fri availabilities
  
  var sheet;
  var rng;
  
  // Clear all interpreter availabilities on sheets Mon-Fri
  for(var i = 0; i < 5; i++) {
    sheet = sched.getSheetByName(week_days[i]);
    
    // Clear all data columns on day sheet
    for(var j = 0; j < 12; j++) {
      rng = cols[j].concat('3:').concat(cols[j]).concat('499');
      sheet.getRange(rng).clearContent();
    }
  }
  
  
}
