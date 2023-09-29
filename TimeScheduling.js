function myFunction() {
    var w2c = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("W2C Availability");
    var availability = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Availabilities");
    var roster = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Roster");
    var tours = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tour Requests");
    var schedule = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tour Schedule");
    var availabilities = availability.getRange('A:GY').getValues();
    var rosters = roster.getRange('A:A').getValues();
  
    var weekNum = schedule.getRange('B1').getValue();
    var weeklySchedule;
  
    if(weekNum == 1) {
      var monday = tours.getRange('D3:E12').getValues();
      var tuesday = tours.getRange('F3:G12').getValues();
      var wednesday = tours.getRange('H3:I12').getValues();
      var thursday = tours.getRange('J3:K12').getValues();
      var friday = tours.getRange('L3:M12').getValues();
  
      // make monday
      var ct = 0
      while(monday[ct][0] != "") {
        var time = Utilities.formatDate(monday[ct][0], "MST", "HH:mm")
        console.log(time)
        var row = 1;  // row in availabilities is row+1
        while(rosters[row][0] != 0) {
          row++;
          
        }
        ct++
      }
  
    }
    else if(weekNum == 2) {
  
    }
    else if(weekNum == 3)  {
  
    }
    else if(weekNum == 4) {
  
    }
    else if(weekNum == 5) {
  
    }
    else {
      // enter a valid week number
    }
  }
  