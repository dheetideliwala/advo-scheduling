// check for saturday sunday tours OR just don't put a number in for that

function scheduleGuides() {
    // get sheet names
    var roster = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Roster")
    var week = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
    sheetName = week.getName()
    var weekNum = 0
    if(sheetName == "Week 1") 
      weekNum = 5
    else if(sheetName == "Week 2")
      weekNum = 6
    else if(sheetName == "Week 3")
      weekNum = 7
    else if(sheetName == "Week 4")
      weekNum = 8
    else if(sheetName == "Week 5")
      weekNum = 9
    console.log(weekNum)
    var tours = week.getRange('A:B').getValues()
    var rosterNames = roster.getRange('A:A').getValues()
    var j = 0
    while(rosterNames[j] != "") {
      j = j+1
    }
    var names = roster.getRange(1, 1, j , 1).getValues()
    var newbie = roster.getRange(1, 2, j, 1).getValues()
    var LOA = roster.getRange(1, 3, j, 1).getValues()
    var sumScheduled = roster.getRange(1, 4, j, 1).getValues()
    var week1 = roster.getRange(1, 5, j, 1).getValues()
    var week2 = roster.getRange(1, 6, j, 1).getValues()
    var week3 = roster.getRange(1, 7, j, 1).getValues()
    var week4 = roster.getRange(1, 8, j, 1).getValues()
    var week5 = roster.getRange(1, 9, j, 1).getValues()
  
    var i = 0
  
    // keep running until current cell and the one after it is empty
    while((tours[i][0] != "") || (tours[i+1][0] != "")) {
      currentCell = tours[i][0].toString()
  
      // find which cells contain tour times with the unique substring of ":", catches dates too because format is in Date format w times HH:MM:SS
      if(currentCell.indexOf(":") > -1 && currentCell.indexOf("Group Visit") == -1 && currentCell.indexOf("Final Tour") == -1) {
  
          // get an updated week range
          if(sheetName == "Week 1") 
            week1 = roster.getRange(1, 5, j, 1).getValues()
          else if(sheetName == "Week 2")
            week2 = roster.getRange(1, 6, j, 1).getValues()
          else if(sheetName == "Week 3")
            week3 = roster.getRange(1, 7, j, 1).getValues()
          else if(sheetName == "Week 4")
            week4 = roster.getRange(1, 8, j, 1).getValues()
          else if(sheetName == "Week 5")
            week5 = roster.getRange(1, 9, j, 1).getValues()
  
          numOfGuides = tours[i][1]
          // checks that the currentCell is not just a day value but a tour time
          if(numOfGuides != "") {
            
            // starting with the first name on the list, check if they are LOA, still a newbie, been selected for this current week already, or already scheduled for 2 tours in the month
            var selected = 0
            var subi = i+1
            while(tours[subi][0] != "" && selected != numOfGuides) { 
                // console.log(tours[subi][0], subi, currentCell)
                name = tours[subi][0].toString()
                // console.log(name)
                // find the name on roster page and row number
                var found = false
                for(var row = 1; row < names.length; row++) {
                  var compareName = names[row].toString()
                  // console.log(name, compareName)
                  if(name == compareName) {
                    found = true;
                    // console.log(name)
                    break
                  }
                }
                // console.log(name)
  
                // if name is found, check if that they are NOT LOA or a newbie or already been scheduled in the week or already been scheduled for 2 tours in the month
                var LOAcheck = true, newbieCheck = true, scheduledInWeekCheck = true, scheduledCheck = true;
                if(found) {
                  // console.log(name)
                  console.log(name, week1[row])
                  if(LOA[row].toString() == "yes")
                    LOAcheck = false;
                  if(newbie[row].toString() == "yes")
                    newbieCheck = false;
                  if(sheetName == "Week 1" && week1[row] != "") 
                    scheduledInWeekCheck = false;
                  if(sheetName == "Week 2" && week2[row] != "")
                    scheduledInWeekCheck = false;
                  if(sheetName == "Week 3" && week3[row] != "")
                    scheduledInWeekCheck = false;
                  if(sheetName == "Week 4" && week4[row] != "")
                    scheduledInWeekCheck = false;
                  if(sheetName == "Week 5" && week5[row] != "")
                    scheduledInWeekCheck = false;
                  if(sumScheduled[row] > 1)
                    scheduledCheck = false;
                  
                  // console.log(name, LOAcheck, newbieCheck, scheduledInWeekCheck, scheduledCheck)
  
                  // if any of these fail, then skip over
                  if(LOAcheck && newbieCheck && scheduledInWeekCheck && scheduledCheck) {
                    // console.log(name)
  
                    // select this guide by adding an X to their name in the Week tab, adding a 1 to the Week they're scheduled for in the Roster tab
                    week.getRange(subi+1, 2, 1, 1).setValue("X")
                    scheduleNum = roster.getRange(row+1, weekNum, 1, 1).getValue()
                    roster.getRange(row+1, weekNum, 1, 1).setValue(scheduleNum + 1)
                    console.log("setting number")
                    selected++;
  
                  }
                }
  
                subi = subi+1
            }
          }
          
      }
  
      i = i+1
    }
    
  }
  
  function newMonth() {
    var roster = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Roster")
    weeks = roster.getRange("E2:I").setValue("")
  }
  