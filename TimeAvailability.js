// function deleteAvailability () {
//   var availability = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Availabilities");
//   var clear = availability.getRange('C3:GY');
//   clear.clearContent();

//   // // clear existing availability
//   // // TAKES A LONG TIME
//   // for(var i = 0; i < 150; i++) {
//   //   for(var j = 0; j < 205; j++) {
//   //     clear[i][j] = "";
//   //   }
//   // }
//   // var range = availability.getRange('C3:GY');
//   // range.setValues(clear)
// }

function addNewAvailability() {
    // get spreadsheet names
    var w2c = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("W2C Availability");
    var availability = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Availabilities");
    var roster = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Roster");
  
    // go through each w2c availability
    // ct is the row # on w2c
    var names = w2c.getRange('A3:A').getValues();
    var days = w2c.getRange('B3:B').getValues();
    var startTimes = w2c.getRange('C3:C').getValues();
    var endTimes = w2c.getRange('D3:D').getValues();
    var loa = w2c.getRange('H3:H').getValues();
    var redo = w2c.getRange('I3:I').getValues();
    var tab = availability.getRange('A:GY');
    var check = availability.getRange('A1:GY2').getValues();
    var tabNames = availability.getRange('A:A').getValues();
    var newbies = w2c.getRange('J3:J').getValues();
  
    var ct = 0; // change back to 0
    nameCheck = (names[ct]).toString(); // first name
    var rowName = 2;
    var first = 0;
  
    // loop to go through each w2c availability
    while(names[ct] && names[ct][0] != "") {
      var day = (days[ct]).toString();
  
      // skip LOA and REDO and Sat/Sun availabilities and newbie
      if(loa[ct][0] != "yes" && redo[ct][0] != "REDO" && day != "Sat" && day != "Sun" && newbies[ct][0] != "yes") {
        var name = (names[ct]).toString();
        
        // find row number for name if first or different name
        if(first == 0) {
          var found = false;
          while(!found) {
            if(tabNames[rowName] == name) {
              found = true;
              //console.log("looking for ", name, "at row ", rowName)
            }
            rowName++;
          }
  
          first = 1;
        }
        else if(nameCheck != name) {
          rowName = 2;
          nameCheck = name;
  
          var found = false;
          while(!found) {
            if(tabNames[rowName] == "") {
              console.log(tabNames[rowName])
              console.log(name, day, "not found")
              ct++;
              name = (names[ct]).toString();
              day = (days[ct]).toString();
              rowName = 1;
            }
            else if(tabNames[rowName] == name) {
              found = true;
            }
            rowName++;
          }
  
        }
    
        // change times before 8:00am to 8:00am
        var startTime = Utilities.formatDate(new Date(startTimes[ct][0]), "MST", "HH:mm");
        var morning = "08:00"
        if(startTime < morning) {
          startTime = morning
        }
        // change times after 6:00pm to 6:00pm
        var endTime = Utilities.formatDate(new Date(new Date(endTimes[ct][0]).getTime() - 60*60*1000), "MST", "HH:mm"); // subtracting an hour for buffer time (make tour creator to base it off of when available to start a tour)
        var evening = "18:00"
        if(endTime > evening) {
          endTime = evening
        }
  
        var found = false;
        var columnNum = 2;
  
        // searches for time frame
        while(!found) {
          var dayCheck = check[0][columnNum];
          var timeCheck = Utilities.formatDate(new Date(check[1][columnNum]), "MST", "HH:mm");
          
          // time frame found, change time frame to all yes
          while(dayCheck == day && timeCheck >= startTime && timeCheck <= endTime) {
            columnNum++
            found = true;
            var cell = availability.getRange(rowName, columnNum)
            cell.setValue("Yes")
            //console.log(name, dayCheck, timeCheck, day, startTime, endTime, rowName)
            // console.log(name, dayCheck, startTime, endTime)
            dayCheck = check[0][columnNum];
            timeCheck = Utilities.formatDate(new Date(check[1][columnNum]), "MST", "HH:mm");
          }
  
          // time frame not found
          if(columnNum == 207) {
            found = true;
          }
          columnNum++;
        }
        
        // // number of columns to write yes in, leaving 1 hour buffer
        // var diff = endTime - startTime;
        // // round up if third decimal is < 5
        // // if((temp * 1000) % 10 < 5) {
        //   diff = Math.ceil((diff-0.04) * 100) / 100;
        //   diff = diff * 100
        // // }
        // // // round down if third decimal is >= 5
        // // else {
        //   // diff = Math.ceil((diff * 100) - 1);
        //   // diff = diff - 4
        // // }
  
        // //find the column
        // var found = false;
        // var columnNum = 2;
        // while(!found) {
        //   var dayCheck = Utilitiescheck[0][columnNum];
        //   var timeCheck = check[1][columnNum];
                  
        //   if(dayCheck == day && timeCheck >= startTime) {
        //     found = true;
        //   }
        //   columnNum++
        // }
  
        // // set yes's
        // for(var i = 0; i < diff; i++) {
        //   var cell = availability.getRange(rowName, columnNum);
        //   cell.setValue("yes");
        //   columnNum++;
        // }
  
      }
      ct++;
    }
  }
  
  function deleteOldAvailability() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Availabilities")
  
    var names = sheet.getRange('A1:A').getValues();
    // console.log(names);
  
    var rowNum = 3;
  
    while(names[rowNum][0] != "") {
      console.log(names[rowNum][0])
      for(var columnNum = 3; columnNum <= 207; columnNum++) {
        var cell = sheet.getRange(rowNum, columnNum)
        cell.setValue("No");
        // console.log(cell);
      }
      rowNum++;
    }
    // var range = sheet.getRange('C3:GY');
    // sheet.setActiveRange(availability);
    // sheet.setActiveRange()
  }
  
  // not being used
  function fillInNo() {
  // get spreadsheet names
    var w2c = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("W2C Availability");
    var availability = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Availabilities");
    var roster = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Roster");
  
    // go through each w2c availability
    // ct is the row # on w2c
    var names = w2c.getRange('A3:A').getValues();
    var days = w2c.getRange('B3:B').getValues();
    var startTimes = w2c.getRange('C3:C').getValues();
    var endTimes = w2c.getRange('D3:D').getValues();
    var loa = w2c.getRange('H3:H').getValues();
    var redo = w2c.getRange('I3:I').getValues();
    var tab = availability.getRange('A:GY');
    var check = availability.getRange('A1:GY2').getValues();
    var tabNames = availability.getRange('A:A').getValues();
    var newbies = w2c.getRange('J3:J').getValues();
  
    var ct = 0; // change back to 0
    nameCheck = (names[ct]).toString(); // first name
    var rowName = 2;
    var first = 0;
  
    // loop to go through each w2c availability
    while(names[ct] && names[ct][0] != "") {
      var day = (days[ct]).toString();
  
      // skip REDO and Sat/Sun availabilities and newbie
      if(loa[ct][0] != "yes" && redo[ct][0] != "REDO" && day != "Sat" && day != "Sun" && newbies[ct][0] != "yes") {
        var name = (names[ct]).toString();
        
        // find row number for name if first or different name
        if(first == 0) {
          var found = false;
          while(!found) {
            if(tabNames[rowName] == name) {
              found = true;
              //console.log("looking for ", name, "at row ", rowName)
            }
            rowName++;
          }
  
          first = 1;
        }
        else if(nameCheck != name) {
          rowName = 2;
          nameCheck = name;
  
          var found = false;
          while(!found) {
            if(tabNames[rowName] == "") {
              console.log(tabNames[rowName])
              console.log(name, day, "not found")
              ct++;
              name = (names[ct]).toString();
              day = (days[ct]).toString();
              rowName = 1;
            }
            else if(tabNames[rowName] == name) {
              found = true;
            }
            rowName++;
          }
  
        }
    
        // change times before 8:00am to 8:00am
        var startTime = Utilities.formatDate(new Date(startTimes[ct][0]), "MST", "HH:mm");
        var morning = "08:00"
        if(startTime < morning) {
          startTime = morning
        }
        // change times after 6:00pm to 6:00pm
        var endTime = Utilities.formatDate(new Date(new Date(endTimes[ct][0]).getTime() - 60*60*1000), "MST", "HH:mm"); // subtracting an hour for buffer time (make tour creator to base it off of when available to start a tour)
        var evening = "18:00"
        if(endTime > evening) {
          endTime = evening
        }
  
        var found = false;
        var columnNum = 2;
  
        // searches for time frame
        while(!found) {
          var dayCheck = check[0][columnNum];
          var timeCheck = Utilities.formatDate(new Date(check[1][columnNum]), "MST", "HH:mm");
          
          // time frame found, change time frame to all yes
          while(dayCheck == day && timeCheck >= startTime && timeCheck <= endTime) {
            columnNum++
            found = true;
            var cell = availability.getRange(rowName, columnNum)
            cell.setValue("No")
            //console.log(name, dayCheck, timeCheck, day, startTime, endTime, rowName)
            // console.log(name, dayCheck, startTime, endTime)
            dayCheck = check[0][columnNum];
            timeCheck = Utilities.formatDate(new Date(check[1][columnNum]), "MST", "HH:mm");
          }
  
          // time frame not found
          if(columnNum == 207) {
            found = true;
          }
          columnNum++;
        }
        
        // // number of columns to write yes in, leaving 1 hour buffer
        // var diff = endTime - startTime;
        // // round up if third decimal is < 5
        // // if((temp * 1000) % 10 < 5) {
        //   diff = Math.ceil((diff-0.04) * 100) / 100;
        //   diff = diff * 100
        // // }
        // // // round down if third decimal is >= 5
        // // else {
        //   // diff = Math.ceil((diff * 100) - 1);
        //   // diff = diff - 4
        // // }
  
        // //find the column
        // var found = false;
        // var columnNum = 2;
        // while(!found) {
        //   var dayCheck = Utilitiescheck[0][columnNum];
        //   var timeCheck = check[1][columnNum];
                  
        //   if(dayCheck == day && timeCheck >= startTime) {
        //     found = true;
        //   }
        //   columnNum++
        // }
  
        // // set yes's
        // for(var i = 0; i < diff; i++) {
        //   var cell = availability.getRange(rowName, columnNum);
        //   cell.setValue("yes");
        //   columnNum++;
        // }
  
      }
      ct++;
    }  
  }
  
  function newAvailabilities() {
    // deleteAvailability();
    // addNewAvailability();
    // fillInNo()
  }
  