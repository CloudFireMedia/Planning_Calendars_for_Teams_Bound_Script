
// Send an email to Team Leaders, soliciting their contribution to the Promotion Planning Calendars for Teams.gsheet
function inviteTeamLeaderstoContributebythefirstFridayOfSeptember() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var staff = getStaff();//get all staff memebers
  var teamLeaders = staff.filter(function(i){return i.isTeamLeader})//remove non-team leaders
  for(var t in teamLeaders){
    var team = teamLeaders[t].team;
    var sheet = ss.getSheetByName(team);

   if(sheet){
     
      
      var to = teamLeaders[t].email;
      var PromotionPlanningCalendarsForTeams_url = SpreadsheetApp.getActiveSpreadsheet().getUrl();
      var teamLeadersContributionDueDate = getFirstFridayOfSeptember();
      var subject = "Please contribute to your team's promotion planning calendar";
      var body = Utilities.formatString("\
Dear %s Leader: <br><br>\
It’s that time again! Please complete the instructions for <a href='%s'>Promotion Planning Calendars for Teams</a>, on behalf of all your team members, by <strong>%s</strong>.<br><br>\
Please reply to this email with any questions.<br><br>\
Thank you!",
                                    team,
                                    PromotionPlanningCalendarsForTeams_url,
                                    teamLeadersContributionDueDate
                                   );
      
      MailApp.sendEmail({
        name     : 'communications@ccnash.org',
        to       : to,
        subject  : subject,
        htmlBody : body
      });
    }//next row
  }//next teamlead    
}



// Send a reminder email 6 days prior to deadline
function remindTeamLeadersToContributeByThisFriday() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var staff = getStaff();//get all staff memebers
  var teamLeaders = staff.filter(function(i){return i.isTeamLeader})//remove non-team leaders
  for(var t in teamLeaders){
    var team = teamLeaders[t].team;
    var sheet = ss.getSheetByName(team);

   if(sheet){
     
      
      var to = teamLeaders[t].email;
      var yourSheetUrl = ss.getUrl() + "#gid=" + sheet.getSheetId();
      var teamLeadersContributionDueDate = getFirstFridayOfSeptember();
      var subject = "Action required! Your team's promotion planning";
      var body = Utilities.formatString("\
Dear %s Leader: <br><br>\
Don't forget, <a href='%s'>your team's Promotion Planning Calendar</a> is due by this coming Friday.<br><br>\
Please let me know if you have any questions or concerns.<br><br>\
Thank you!",
                                    team,
                                    yourSheetUrl,
                                    teamLeadersContributionDueDate
                                   );
      
      MailApp.sendEmail({
        name     : 'communications@ccnash.org',
        to       : to,
        subject  : subject,
        htmlBody : body
      });
    }//next row
  }//next teamlead    
}


function getStaff() {
 
  var staffDataId = '1iiFmdqUd-CoWtUjZxVgGcNb74dPVh-l5kuU_G5mmiHI'; // "Staff Data.gsheet"
  var staffData = SpreadsheetApp.openById(staffDataId);
  var sheet = SpreadsheetApp.openById(staffDataId).getActiveSheet();
  var values = sheet.getDataRange().getValues();
  values = values.slice(sheet.getFrozenRows());//remove headers if any
  
  var staff = values.map(function(c,i,a){
    return {
      name         : [c[0], c[1]].join(' '),
      email        : c[8],
      team         : c[11],
      isTeamLeader : (c[12].toLowerCase()=='yes'),
      jobTitle     : c[4],
    };
  },[]);

  return staff;
}

function getTeamLeader(team) {
  
  var staffDataId = '1iiFmdqUd-CoWtUjZxVgGcNb74dPVh-l5kuU_G5mmiHI'; // "Staff Data.gsheet"
  var staffData = SpreadsheetApp.openById(staffDataId);
  var sheet = SpreadsheetApp.openById(config.files.staffData);
  var values = sheet.getDataRange().getValues();
  for(var v=sheet.getFrozenRows(); v<values1.length; v++)
    if(team == values[v][11] && values[v][12] == "Yes")
        return values[v][8] + ',' + values[v][0] + " " + values[v][1];//email, fl

  
/*
* Get the staff data
*/
  var staffDataRange = staffData.getDataRange();
  var staff = getStaff();
  
  //remove non-team leaders
  var teamLeaders = staff.filter(function(i) {
    return i.isTeamLeader;
  })
  
  var teams = teamLeaders.reduce(function (out, cur) {
    if (cur.isTeamLeader) {    
      out[cur.team] = {
        name:cur.name,
        email:cur.email
      };
    }
    return out;
  }, {});
  
}


/*
* Get the first Friday of September for the current year.
*/
function getFirstFridayOfSeptember() {
  var year = new Date().getFullYear();
  var date = new Date("September 1 " + year);
  var target = 5; // Friday
  if(target != 5) {
    var days = ( 30 - ( target - date.getDay() ) % 7 );
    var time = date.getTime() - ( days * 86400000 );
    date.setTime(time);
  }
  var options = { weekday: 'long', month: 'long', day: 'numeric' };
  teamLeadersContributionDueDate = date.toLocaleString('en-us', options);
  teamLeadersContributionDueDate = 'Friday, ' + teamLeadersContributionDueDate.split(',')[0];
  return teamLeadersContributionDueDate;
}
