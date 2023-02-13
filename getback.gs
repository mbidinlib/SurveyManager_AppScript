/************************************************************
Name:          Getting Back
Purpose:       Resets MyRa to Getting Started page
Author :       Mathew Bidinlib
email  :       mbidinlib@poverty-action.org
Date created:  Jan 20 2020
Date Modified: March 10 2020
Copyright    : Innovations for Poverty action @2020
***********************************************************/

function getback(){
   // Message Box to confirm deletion of sheets and returning to getting started 
    var action = Browser.msgBox("Restart Alert!","Are you sure you want to restart? Continuing will delete all sheets and progress", Browser.Buttons.OK_CANCEL);
  
  if (action == "ok"){ 
    // Define sheet names
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var startpage = ss.getSheetByName("Getting Started")
    var master = ss.getSheetByName("master");
    var gantt = ss.getSheetByName("gantt");
    var project_open = ss.getSheetByName("project_open");
    var project_close = ss.getSheetByName("project_close");
    var task_sumary = ss.getSheetByName("task_summary");
    
    // Make Getting Started sheet active
    startpage.showSheet();
    ss.setActiveSheet(startpage);
        
    // Hide Sheets
    master.hideSheet();
    gantt.hideSheet();
    task_sumary.hideSheet();
    
    var allsheets = ss.getSheets();
    
    for (var s in allsheets){
      var sheet = allsheets[s]
      var name = sheet.getSheetName();
    
    // Loops through project sheets
      if (sheet.isSheetHidden()!= true && name != "master" && name != "gantt" 
          && name != "health_check" && name !=  "task_summary" && name !=  "Getting Started") {   
          var del_sht = ss.getSheetByName(name) 
          ss.deleteSheet(del_sht)    
      }
    }   
    // Format master sheet 
    master.getRange("E12:E13").setBackground("");
    master.getRange("E12:F13").setValue("");
    master.getRange("E12:F13").setBorder(false, false, false, false, false,false);
    master.getRange("E12:F12").setBorder(true, false, false, false, false, false);
  
  }
    
}
