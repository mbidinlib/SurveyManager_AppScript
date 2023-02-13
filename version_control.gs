/*************************************************************
Name:          Automation
Purpose:       Triggers automatic emails to Users/RQ/Supervisors
Author :       Mathew Bidinlib
email  :       mbidinlib@poverty-action.org
Date created:  Jan 20 2020
Date Modified: March 10 2020
Copyright    : Innovations for Poverty action @2020
*************************************************************/


function onEdit(e) {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var gant = ss.getSheetByName("gantt");
  var master = ss.getSheetByName("master");
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var sheetname = sheet.getSheetName();
  var name = sheetname.split("_")[1];
  var team = master.getRange(16,6,10,1).getValues().join(";");
  
  
  //adding new line at the top
  if (name == "tasks" | name == "survey" | name == "impl" | name =="open" | name == "close") {
      var editcell = sheet.getActiveCell().getValue();
      var editcol = sheet.getActiveCell().getColumn();
      var editrow = sheet.getActiveCell().getRow();
      
      if (editcol == 1 && editcell == "New"){
          sheet.insertRowsBefore(editrow, 1);    // inserts a row before
          
          // remove new values and leave blank
          sheet.getRange(editrow+1, editcol).setValue("");
        
          //Add delete validation for new items added
          sheet.getRange(editrow, editcol).setDataValidation(null);
          var valrange = sheet.getRange(editrow, editcol);
          var rule = SpreadsheetApp.newDataValidation().requireValueInList(["Completed","New","Delete"]).build();
          valrange.setDataValidation(rule);   
      } 
     
      // Delete row that is marked as delete    
      if (editcol == 1 && editcell == "Delete"){
            sheet.deleteRow(editrow);
        } 
      
    if (editcol == 1 && editcell == ""){
           var completed_range =  sheet.getRange(editrow,editcol+1)
            completed_range.setFontColor("Black");
    }
        // Change fontColor of completed tasks    
      if (editcol == 1 && editcell == "Completed"){
           var completed_range =  sheet.getRange(editrow,editcol+1)
            completed_range.setFontColor("Gray");
            var completed_task = completed_range.getValue();
            var country_name = master.getRange("C2").getValue();
        
             // Trigger an email to RQ Team (Ghana)
            if(country_name == "Ghana"){
                var project_name = master.getRange("C3").getValue();
                var rq_email = master.getRange("F12").getValue();
                var link = SpreadsheetApp.getActiveSpreadsheet().getUrl();
                var team = master.getRange(16,6,10,1).getValues();
                              
              MailApp.sendEmail(
                rq_email,
                "MyRA_Updates_"+ project_name ,
                "Hello RQ Team,<p>There has been a recent MyRA update for ..." +            
                 project_name + "...Project.</p> <p><b>Task:</b> " + completed_task + "</p> <p><b>Action:</b> Completed </p>" +
                "<b>MyRA Link:</b> " + link + "<p>Thank you</p>",
                 {htmlBody: 
                 "Hello RQ Team,<p>There has been a recent MyRA update for ..." +            
                 project_name + "...Project.</p> <p><b>Task:</b> " + completed_task + "</p> <p><b>Action:</b> Completed </p>" +
                "<b>MyRA Link:</b> " + link + "<p>Thank you</p>"
                }); 
              
            }     
      }
   }
  
}



// Send Reminders on overdue tasks

function notifyOverdue() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName("master");
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var emails = master.getRange(16,6,10,1).getValues();
  var n_emails = emails.filter(String).length;
  var team = master.getRange(16,6,n_emails,1).getValues().join(";");
  var sheets = ss.getSheets();
  var rq_email = master.getRange("F12").getValue();
  var project_name = master.getRange("C3").getValue();
  var link = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  
  for (var i = 0; i < sheets.length; i++) {
    var shtname =  sheets[i].getSheetName();
    var name = shtname.split("_")[1];
    Logger.log(name)
    
    
    if(name == "tasks" | name == "survey" | name == "impl" | name =="open" | name == "close") {
     var sheet = ss.getSheetByName(shtname);
     var range = sheet.getRange(1,5,500,1).getValues();
     //Logger.log(name + "Yes" + sheet + range) 
     var toval = range.filter(String).length;
      
      for(var j = 1; j <= toval; j++) {
        
        if(range[j-1] == "Overdue"){
          
           MailApp.sendEmail(
                rq_email,
                "MyRA_Task_Overdue_"+ project_name ,
                "Hello" + project_name +" Team,<p> A task is overdue and requires your attention</p>" +
                "<p><b>Task:</b>" + sheet.getRange(j,2).getValue() + "</p>",
                 {htmlBody: 
                "Hello" + project_name +"Team,<p> An event is overdue and requires your attention</p>" +
                "<p> <b>Task:</b>" + sheet.getRange(j,2).getValue() + "</p>" +
                 "<b>MyRA Link:</b> " + link + "<p>Thank you</p>"
                 }); 
         
        }
        
      }
                 
    }
       

  }

}