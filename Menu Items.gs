function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('CloudFire')
      .addItem('Update Instructions and Promotion Planning Calendars for Teams', 'setDateValidation')
      .addSeparator()
      .addItem('Nudge Team Leaders Who Haven\'t Contributed After a Week', 'nudgeTeamLeader')
      .addSeparator()
      .addItem("Invite Team Leaders to Contribute by the first Friday of September", "inviteTeamLeaderstoContributebythefirstFridayOfSeptember")
      .addSeparator()
      .addItem("Remind Team Leaders to Contribute 'by this coming Friday'", "remindTeamLeadersToContributeByThisFriday")
      .addToUi();
}