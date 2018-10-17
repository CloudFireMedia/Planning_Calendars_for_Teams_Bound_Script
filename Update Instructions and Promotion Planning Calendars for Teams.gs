
/*
* Updates the 'Promotion Planning Calendars for Teams.gsheet' with instructions 
* and team information from the ' Promotion Planning Calendar for Teams TEMPLATE.gsheet'.
* It also changes the date validation for Col B to 'next year' and protects the ranges 
* of the current year's event data on each team sheet.

***************************************************************************************

Redevelopment Notes from Chad: 

...

***************************************************************************************
*/

/*
* Will change the date validation for Col B in the Template to {next year}.
*/
function setDateValidation () {
  var sheet = SpreadsheetApp.openById("1dXQK6zHEjAwwk9bQibSsQzlpgFRNPtJ_HJY414bAZKA").getSheetByName('TEAM');
  var this_year = new Date().getYear();
  var next_year = new Date().getYear() + 1;
  // Change existing data validation rules that require a date in this_year to require a date in next_year.
  var oldDates = [new Date(this_year,0,1), new Date(this_year,11,31)];
  var newDates = [new Date(next_year,0,1), new Date(next_year,11,31)];    
  var range = sheet.getRange("B3:B50");
  var rules = range.getDataValidations();
  for (var i = 0; i < rules.length; i++) {
    for (var i = 0; i < rules.length; i++) {
      for (var j = 0; j < rules[i].length; j++) {
        var rule = rules[i][j];
        if (rule != null) {
          var criteria = rule.getCriteriaType();
          var args = rule.getCriteriaValues();
          if (criteria == SpreadsheetApp.DataValidationCriteria.DATE_BETWEEN
              && args[0].getTime() == oldDates[0].getTime()
          && args[1].getTime() == oldDates[1].getTime()) {
            // Create a builder from the existing rule, then change the dates.
            rules[i][j] = rule.copy().withCriteria(criteria, newDates).build();
          }
        }
      }
    }
  }
  range.setDataValidations(rules);
  updateInstructionsAndPromotionPlanningCalendarsForTeams()
}

function updateInstructionsAndPromotionPlanningCalendarsForTeams() {
  
  /*
  * Will verify manual execution is intentional and not too early.
  */ 
  var title = 'Warning';
  var prompt = Utilities.formatString(
    "\
Executing this script will result in the deletion of all present team sheets, \n\
including the loss of any data that may have been contributed manually to these sheets.\n\
\n\
Do you wish to proceed?\
")
  if(SpreadsheetApp.getUi().alert(title, prompt, Browser.Buttons.YES_NO) != 'YES') return;  
  
  //  /*
  //  * Will notify Chad that the component is updated and 
  //  * aready for Team Leaders' contributions.
  //  */
  //  var PromotionPlanningCalendarsforTeams_url = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  //  var subject = "Notice: Promotion Planning Calendars for Teams have been updated";
  //  var body = Utilities.formatString("\
  //Communications Director: <br><br>\
  //The 'Instructions' and 'Team' tabs in the <a href='%s'>Promotion Planning Calendars for Teams</a> component have been automatically updated for the current calendar year.<br><br>\
  //Team Leaders will be invited to comment on their Planning Calendars beginning tomorrow.<br><br>\
  //The deadline for contributions is the first Friday of September.\
  //",
  //                                    PromotionPlanningCalendarsforTeams_url
  //                                   );
  //      
  //    MailApp.sendEmail({
  //      name: 'communications@ccnash.org',
  //      to: 'chbarlow@gmail.com',
  //      subject: subject,
  //      htmlBody: body
  //    });
  
  var PromotionPlanningCalendarsForTeamsTEMPLATE_id = "1dXQK6zHEjAwwk9bQibSsQzlpgFRNPtJ_HJY414bAZKA"; // " Promotion Planning Calendars for Teams TEMPLATE.gsheet
  var PromotionPlanningCalendarForTeams_id = '1ZlB3EOL4smtkNhpsvjUxSRI1cXBPuuucbOil49uL2TQ'; // "Promotion Planning Calendars for Teams - Test Copy.gsheet"
  var StaffSheet_id = '1iiFmdqUd-CoWtUjZxVgGcNb74dPVh-l5kuU_G5mmiHI'; // "Staff Data.gsheet"
  
  var PromotionPlanningCalendarsForTeamsTEMPLATE = SpreadsheetApp.openById(PromotionPlanningCalendarsForTeamsTEMPLATE_id);
  var PromotionPlanningCalendarForTeams = SpreadsheetApp.openById(PromotionPlanningCalendarForTeams_id);
  var StaffSheet = SpreadsheetApp.openById(StaffSheet_id);
  
  /*
  * Will copy the information from the template
  * Replace the date variables and save the instructions with the
  * new date information on the target sheet.
  */
  var instructionsTemplate_F4 = PromotionPlanningCalendarsForTeamsTEMPLATE.getSheetByName('Instructions').getRange("F4");
  var instructionsTemplate_D7 = PromotionPlanningCalendarsForTeamsTEMPLATE.getSheetByName('Instructions').getRange("D7");
  var instructionsTemplate_D8 = PromotionPlanningCalendarsForTeamsTEMPLATE.getSheetByName('Instructions').getRange("D8");
  var instructionsTemplate_D9 = PromotionPlanningCalendarsForTeamsTEMPLATE.getSheetByName('Instructions').getRange("D9");
  
  var instructions_F4 = PromotionPlanningCalendarForTeams.getSheetByName('Instructions').getRange("F4");
  var instructions_D7 = PromotionPlanningCalendarForTeams.getSheetByName('Instructions').getRange("D7");
  var instructions_D8 = PromotionPlanningCalendarForTeams.getSheetByName('Instructions').getRange("D8");
  var instructions_D9 = PromotionPlanningCalendarForTeams.getSheetByName('Instructions').getRange("D9");
  
  var teamLeadersContributionDueDate = getFirstWeek_();
  
  var currentYear = new Date().getFullYear();
  var instructionDetails_F4 = instructionsTemplate_F4.getValue().toString();
  instructionDetails_F4 = instructionDetails_F4.replace(/TEAM_LEADERS_CONTRIBUTION_DUE_DATE/g, teamLeadersContributionDueDate)
  
  var instructionDetails_D7 = instructionsTemplate_D7.getValue().toString();
  instructionDetails_D7 = instructionDetails_D7.replace(/UPCOMING_YEAR/g, currentYear + 1);
  
  var instructionDetails_D8 = instructionsTemplate_D8.getValue().toString();
  instructionDetails_D8 = instructionDetails_D8.replace(/UPCOMING_YEAR/g, currentYear + 1);
  
  var instructionDetails_D9 = instructionsTemplate_D9.getValue().toString();
  instructionDetails_D9 = instructionDetails_D9.replace(/UPCOMING_YEAR/g, currentYear + 1);
  
  instructions_F4.setValue(instructionDetails_F4);
  instructions_D7.setValue(instructionDetails_D7);
  instructions_D8.setValue(instructionDetails_D8);
  instructions_D9.setValue(instructionDetails_D9);
  Logger.log("Instructions have been updated");
  
  // Get a list of team names 
  var TEAM_COLUMN_INDEX = 12;
  var EMPLOYEE_FIRST_RECORD = 3;
  
  var teamArray = [];
  var staff = StaffSheet.getSheetByName("Staff");
  
  for(var i = EMPLOYEE_FIRST_RECORD; i<staff.getLastRow(); i++) {
    var teamName = staff.getRange(i, TEAM_COLUMN_INDEX).getValue();
    if(teamName == '') teamName = 'No Team';
    if(teamArray.indexOf(teamName) == -1) {
      teamArray.push(teamName);
    }
  }
  var teams = teamArray.sort(); 
  
  
  /*
  * Remove the existing team sheets.
  */ 
  var sheets = PromotionPlanningCalendarForTeams.getSheets();
  for(var i = 1; i < sheets.length; i++)
  {
    var sheetName = sheets[i].getName();
    if(sheetName.toLowerCase().indexOf('team')) PromotionPlanningCalendarForTeams.deleteSheet(sheets[i]);
  }
  
  // Enumerate the team names and add new sheets.
  for(var teamIndex = 0; teamIndex < teams.length; teamIndex++ ) {
    var teamName = teams[teamIndex];
    
    var teamTemplate = PromotionPlanningCalendarsForTeamsTEMPLATE.getSheetByName("TEAM");
    var newTeamSheet = teamTemplate.copyTo(PromotionPlanningCalendarForTeams);
    
    // Copy formulas and replace TEAM in the formula with the team name from the calendar.
    // Doesn't work well for the "N/A" and "No Team" sheets.
    var teamRanges = teamTemplate.getRange( 3, 1, teamTemplate.getLastColumn(), teamTemplate.getLastColumn() );
    var formulas = teamRanges.getFormulas();
    for (var i in formulas) {
      for (var j in formulas[i]) {
        var templateFormula = formulas[i][j];
        formulas[i][j] = templateFormula.replace(/TEAM/g, teamName);
        Logger.log(formulas[i][j]);
      }
    }
    Logger.log('Formulas replaced with team names.');
    
    Logger.log('Setting templateFormulas to: ' + formulas);
    var newTeamRanges = newTeamSheet.getRange( 3, 1, teamTemplate.getLastColumn(), teamTemplate.getLastColumn() );
    newTeamRanges.setFormulas(formulas);
    
    // Get values
    var sheetTitle = newTeamSheet.getRange("A1");
    var suggestedUpcomingTier = newTeamSheet.getRange("A2");
    var suggestedUpcomingYear = newTeamSheet.getRange("B2");
    var upcomingPromotionLevel = newTeamSheet.getRange("C2");
    var currentDateYear = newTeamSheet.getRange("D2");
    
    // Set the sheets properties and values
    
    var currentYear = new Date().getFullYear();
    newTeamSheet.setName(teamName);
    sheetTitle.setValue( sheetTitle.getValue().toString().replace(/TEAM/g, teamName) );
    sheetTitle.setValue( sheetTitle.getValue().toString().replace(/UPCOMING_YEAR/g, currentYear + 1) );
    suggestedUpcomingTier.setValue( suggestedUpcomingTier.getValue().toString().replace(/UPCOMING_YEAR/g, currentYear + 1) );
    suggestedUpcomingYear.setValue( suggestedUpcomingYear.getValue().toString().replace(/UPCOMING_YEAR/g, currentYear + 1) );
    upcomingPromotionLevel.setValue( upcomingPromotionLevel.getValue().toString().replace(/CURRENT_YEAR/g, currentYear) );
    currentDateYear.setValue( currentDateYear.getValue().toString().replace(/CURRENT_YEAR/g, currentYear) );
    console.log("Team Sheets have been updated");
  }
  protectRanges();
}

function protectRanges() {
  
  var currentYear = new Date().getFullYear();
  
  var sheets = SpreadsheetApp.getActive().getSheets();  
  var me = Session.getEffectiveUser();
  
  for (var sheetIndex = 1; sheetIndex < sheets.length; sheetIndex++) { 
    var sheet = sheets[sheetIndex];
    
    var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    for (var i = 0; i < protections.length; i++) {
      var protection = protections[i];
      if (protection.canEdit()) {
        protection.remove();
      }
    };
    
    var startRow = 3;
    var startColumn = 3;
    var numberOfRows = sheet.getLastRow() - 2;
    var numberOfColumns = 3;
    var range = sheet.getRange(startRow, startColumn, numberOfRows, numberOfColumns);
    var border_range = sheet.getRange(startRow, 1, numberOfRows, 6);
    border_range.setBorder(null, null, true, null, null, null, '#163c47', SpreadsheetApp.BorderStyle.DASHED);  // top left bottom right vert horz color type
    var protection = range.protect().setDescription(sheet.getName() + ' ' + (currentYear) + ' ' + 'Events');
    
    // Ensure the current user is an editor before removing others. Otherwise, if the user's edit
    // permission comes from a group, the script throws an exception upon removing the group
    protection.removeEditors(protection.getEditors());
    if (protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    }
  }
} 




function getWeeksInMonth_(month, year){
  var year = new Date().getYear();
  var month = 9;
  var weeks_array=[],
      firstDate=new Date(year, month, 1),
      lastDate=new Date(year, month+1, 0), 
      numDays= lastDate.getDate();
  
  var start=1;
  var end=7-firstDate.getDay();
  while(start<=numDays){
    weeks_array.push({start:start,end:end});
    start = end + 1;
    end = end + 7;
    end = start === 1 && end === 8 ? 1 : end;
    if(end>numDays)
      end=numDays;    
  }
  return weeks_array;
}

function nth_(ordinand) {
  if(ordinand>3 && ordinand<21) return 'th'; 
  switch (ordinand % 10) {
    case 1:  return "st";
    case 2:  return "nd";
    case 3:  return "rd";
    default: return "th";
  }
} 

function getFirstWeek_() {
  var weeks_array = getWeeksInMonth_();
  var year = new Date().getYear();
  var month = 9;
  
  var m_names = ['January', 'February', 'March', 
                 'April', 'May', 'June', 'July', 
                 'August', 'September', 'October', 'November', 'December'];
  d = new Date(year, month, 1, 0, 0, 0, 0);
  var n = m_names[d.getMonth()];
  
  for (var weeks_array_index = 0; weeks_array_index < weeks_array.length; weeks_array_index++) {
    var current_start = weeks_array[weeks_array_index].start; 
    var current_end = weeks_array[weeks_array_index].end;
    var first_Friday = current_end - 1;
    var teamLeadersContributionDueDate = "Friday, " + n + " " + first_Friday + nth_(current_end);
    
    if (current_end - current_start == 6 ) { 
            
     
       Logger.log("Friday, " + n + " " + first_Friday + nth_(current_end));
       return teamLeadersContributionDueDate;
    }
  }
}

