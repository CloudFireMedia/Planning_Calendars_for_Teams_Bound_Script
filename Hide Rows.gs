function onEdit(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var cell = ss.getActiveCell()
  var cellValue = cell.getValue();
  var next_year = new Date().getYear() + 1;
  if(cell.getColumn() == 6)
    if(cellValue === "This event will not recur next year") { // cellValue === "This event belongs to another team" || 
      activeSheet.hideRow(cell);
      ui.alert(
      'Hey thanks!',
      'Thanks for indicating that this will NOT recur in the next calendar next year.\n\n\
      In order to simplify your calendar, the event has been automtically hidden.\n\n\
      If needed, you can unhide it by clicking on the up/down arrows on the row number\n\
      column on the very left-hand side of your screen.\n\n'
      ,ui.ButtonSet.OK);
    }
  }
