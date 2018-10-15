/* 
* Send a reminder email to Team Leaders who have not manually entered data to 
* their Team's Planning Calendar.
* 
* Bug: the script WILL NOT send an email to any team leader whose planning calendar
* contains a {next year} date in Col D, EVEN IF IT THAT DATE WAS IMPORTED from the 
* Promotion Deadlines Calendar and not entered manually by the staff person. Whatevs.
*/

function getStaffArray() {
  var sheet = SpreadsheetApp.openById('1HEOWmNPo32uhR6N1XkviYiDM7KdAnaYycKDH9fz3OXE'); // "Staff Data. - Test
  var values = sheet.getDataRange().getValues();
  values = values.slice(sheet.getFrozenRows());//remove headers if any
  var staff_array = values.map(function(c,i,a){
    return {
      current_value_is_staff_first_name  : c[0],
      current_value_is_staff_email       : c[8],
      current_value_is_team_name         : c[11],
      current_value_is_team_leader       : (c[12].toLowerCase()==='yes'),
    };
  },[]);   
  return staff_array;
}

function nudgeTeamLeader() {
  var staff_array = getStaffArray();  //get all staff members
  var team_leaders_array = staff_array.filter(function(i){return i.current_value_is_team_leader}); //remove non-team leaders 
  for (var team_leaders_array_index = 0; team_leaders_array_index < team_leaders_array.length; team_leaders_array_index++) {  
    var Spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); 
    var current_team_name = team_leaders_array[team_leaders_array_index].current_value_is_team_name; 
    var current_team_sheet = Spreadsheet.getSheetByName(current_team_name);
    var current_leader_first_name = team_leaders_array[team_leaders_array_index].current_value_is_staff_first_name; 
    var current_leader_email = team_leaders_array[team_leaders_array_index].current_value_is_staff_email;
    var current_leader_sheet_url = Spreadsheet.getUrl() + "#gid=" + current_team_sheet.getSheetId();     
    var today_date = new Date(); //Wed Oct 10 2018 08:49:53 GMT-0500 (CDT)
    var this_year = today_date.getFullYear(); //2018.0
    var current_team_sheet_dates_array = current_team_sheet.getRange("D3:D50").getValues();   
    for (var current_team_sheet_dates_array_index = 0; current_team_sheet_dates_array_index < current_team_sheet_dates_array.length; current_team_sheet_dates_array_index++){
      var current_year_for_dates_array_object = new Date(current_team_sheet_dates_array[current_team_sheet_dates_array_index][0]).getYear();
      if (current_year_for_dates_array_object !== this_year) 
        
        break; }
    
    var last_row = current_team_sheet.getMaxRows();
    var first_unprotected_row = current_team_sheet_dates_array_index + 3;
    if (current_team_sheet.getRange("C" + first_unprotected_row + ":E" + last_row).isBlank() && current_team_sheet.getRange("A3:B50").isBlank() && current_team_sheet.getRange("F3:F50").isBlank()) {   
      var subject = "Just checking in";
      var body = Utilities.formatString("\
%s, <br><br>\
I noticed you haven't had a chance to make any changes or additions to <a href='%s'>your team's Promotion Planning Calendar</a> yet.<br><br>\
Is there anything I can do to help?<br><br>\
",
                                        current_leader_first_name,
                                        current_leader_sheet_url
                                       );    
      MailApp.sendEmail({
        name     : 'communications@ccnash.org',
        to       : current_leader_email,
        subject  : subject,
        htmlBody : body
      });    
    }
  } 
}
