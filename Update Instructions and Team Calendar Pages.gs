
/*
* Updates the 'Promotion Planning Calendars for Teams.gsheet' with instructions 
* and team information from the ' Promotion Planning Calendar for Teams TEMPLATE.gsheet'.
* It also protects the ranges of the current year's event data.
*/

/*
******************************************************************************************************************************

Redevelopment Notes from Chad: 

This script does great, but it would be better if the formatting from the 'Instructions' tab on the TEMPLATE could be carried over
to the destination sheet

******************************************************************************************************************************
*/

function updateInstructionsAndPromotionPlanningCalendarsForTeams() {
  
  /*
  * Will verify manual execution is intentional and not too  
  * early.
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
  var PromotionPlanningCalendarForTeams_id = '1JSyVOmWqqnJrAfs_eBHd2MW1SdtZImzJdRJAsW0aPA4'; // "Promotion Planning Calendars for Teams - Test Copy.gsheet"
  var StaffSheet_id = '1HEOWmNPo32uhR6N1XkviYiDM7KdAnaYycKDH9fz3OXE'; // "Staff Data.gsheet"
  
  var PromotionPlanningCalendarsForTeamsTEMPLATE = SpreadsheetApp.openById(PromotionPlanningCalendarsForTeamsTEMPLATE_id);
  var PromotionPlanningCalendarForTeams = SpreadsheetApp.openById(PromotionPlanningCalendarForTeams_id);
  var StaffSheet = SpreadsheetApp.openById(StaffSheet_id);
  
  /*
  * Will copy the information from the template
  * Replace the date variables and save the instructions with the
  * new date information on the respected sheet.
  */
  var instructionsTemplate = PromotionPlanningCalendarsForTeamsTEMPLATE.getSheetByName('Instructions').getRange("C4");
  var instructions = PromotionPlanningCalendarForTeams.getSheetByName('Instructions').getRange("C4");
  
  var firstFridayOfSeptember = getFirstFridayOfSeptember_();
  
  var currentYear = new Date().getFullYear();
  var instructionDetails = instructionsTemplate.getValue().toString();
  instructionDetails = instructionDetails.replace(/FIRST_FRIDAY_OF_SEPTEMBER_OF_THE_CURRENT_YEAR/g, firstFridayOfSeptember)
  .replace(/CURRENT_YEAR/g, currentYear)
  .replace(/UPCOMING_YEAR/g, currentYear + 1);
  
  instructions.setValue(instructionDetails);
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
  };
  protectRanges();
};

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
    var protection = range.protect().setDescription(sheet.getName() + ' ' + (currentYear) + ' ' + 'Events');
    
    // Ensure the current user is an editor before removing others. Otherwise, if the user's edit
    // permission comes from a group, the script throws an exception upon removing the group
    protection.addEditor(me);
    protection.removeEditors(protection.getEditors());
    if (protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    }
  }
} // protectRanges()  

function getFirstFridayOfSeptember_() {
  var year = new Date().getFullYear();
  var date = new Date("September 1 " + year);
  var target = 5; // Friday
  if(target != 5) {
    var days = ( 30 - ( target - date.getDay() ) % 7 );
    var time = date.getTime() - ( days * 86400000 );
    
    // setting full timestamp here
    date.setTime(time);
  }
  var options = { weekday: 'long', month: 'long', day: 'numeric' };
  teamLeadersContributionDueDate = date.toLocaleString('en-us', options);
  teamLeadersContributionDueDate = 'Friday, ' + teamLeadersContributionDueDate.split(',')[0];
  return teamLeadersContributionDueDate;
}