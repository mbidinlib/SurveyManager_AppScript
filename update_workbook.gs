/************************************************************
Name:          UpdateWorksheet
Purpose:       Updates the Gannt chart and the task summary sheets
Author :       Mathew Bidinlib
email  :       mbidinlib@poverty-action.org
Date created:  Jan 20 2020
Date Modified: March 10 2020
Copyright    : Innovations for Poverty action @2020
***********************************************************/


function update_workbook() {
  var ss = SpreadsheetApp.getActive();
  var allsheets = ss.getSheets();
  var gantt = ss.getSheetByName("gantt");
  var freq = ss.getSheetByName("master").getRange("C11").getValue(); 
  var newval = - 8
  var newval1 = -20
  var newval2 = -3 
  
  for (var j = 8; j <= 104; j++){
   
    //Daily Dates
    if (freq == "Daily"){
      var newval = newval + 1
      gantt.getRange(1, j).setFormula("=TODAY() + " + newval ); 
    }
    
    //Weekly Dates
     if (freq == "Weekly"){
      var newval1 = newval1 + 7
      gantt.getRange(1, j).setFormula("=TODAY() - WEEKDAY(TODAY(), 2) + " + newval1 );
    }
    
    //Monthly Dates
    
    if (freq == "Monthly"){
      var newval2 = newval2 +1
      var mon = gantt.getRange("H1").getValue();
      gantt.getRange(1, j).setFormula("=EOMONTH(TODAY()," + newval2 + ") + 1" );    
    }      
  }
 
  
  // ****** TASK SUMMATY **** \\
  
  var task_sum = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("task_summary");  
  task_sum.getRange(3, 2,500,49).setValue("");  
 
  for (var s in allsheets){
    var sheet = allsheets[s]
    var name = sheet.getSheetName();
    
    // Loops through project sheets
    if (sheet.isSheetHidden()!= true && name != "master" && name != "gantt" 
        && name != "health_check" && name !=  "task_summary") {
      var sht_name = sheet.getSheetName();  
      var sheet_n = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sht_name);
   
        //**Task Summary Overdue 
      var array = task_sum.getRange("B3:B1000").getValues();
      var start = array.filter(String).length + 3; // removes the empty values and counts the length
      sheet_n.getRange('E:E').createFilter();
      var criteria = SpreadsheetApp.newFilterCriteria()
      .setHiddenValues(['Complete','Ongoing', 'Scheduled', 'Upcoming', 'Not Scheduled']).build();    
      sheet_n.getFilter().setColumnFilterCriteria(5, criteria);
      //Populate Overdue columns
      sheet_n.getRange("B3:B140").copyTo(task_sum.getRange(start, 2, 140,1), {contentsOnly:true});
      sheet_n.getRange("D3:D140").copyTo(task_sum.getRange(start, 3, 140,1), {contentsOnly:true});
      sheet_n.getRange("F3:G140").copyTo(task_sum.getRange(start, 4, 140,2), {contentsOnly:true});
      sheet_n.getRange("H3:I140").copyTo(task_sum.getRange(start, 6, 140,2), {contentsOnly:true});
      var array = task_sum.getRange("B3:B1000").getValues();
      var start = array.filter(String).length + 3; // removes the empty values and counts the length 
      task_sum.getRange(start + 3,8,140,1).setValue("");    
      
      
      //**Task Summary Ongoing  
      var array = task_sum.getRange("J3:J1000").getValues();
      var start = array.filter(String).length + 3; // removes the empty values and counts the length
      var criteria = SpreadsheetApp.newFilterCriteria()
      .setHiddenValues(['Complete','Overdue', 'Scheduled', 'Upcoming', 'Not Scheduled']).build();    
      sheet_n.getFilter().setColumnFilterCriteria(5, criteria);      
      //Populate Ongoind columns
      sheet_n.getRange("B3:B140").copyTo(task_sum.getRange(start, 10, 140,1), {contentsOnly:true});
      sheet_n.getRange("D3:D140").copyTo(task_sum.getRange(start, 11, 140,1), {contentsOnly:true});
      sheet_n.getRange("F3:G140").copyTo(task_sum.getRange(start, 12, 140,2), {contentsOnly:true});
      sheet_n.getRange("H3:I140").copyTo(task_sum.getRange(start, 14, 140,2), {contentsOnly:true});
      var array = task_sum.getRange("J3:J1000").getValues();
      var start = array.filter(String).length + 3; // removes the empty values and counts the length 
      task_sum.getRange(start + 3,8,140,1).setValue("");    
      
       //** Task Summary Upcoming
      var array = task_sum.getRange("R3:R1000").getValues();
      var start = array.filter(String).length + 3; // removes the empty values and counts the length
      var criteria = SpreadsheetApp.newFilterCriteria()
      .setHiddenValues(['Complete','Overdue', 'Scheduled', 'Ongoing', 'Not Scheduled']).build();    
      sheet_n.getFilter().setColumnFilterCriteria(5, criteria);
      //Populate Upcoming Columns
      sheet_n.getRange("B3:B140").copyTo(task_sum.getRange(start, 18, 140,1), {contentsOnly:true});
      sheet_n.getRange("D3:D140").copyTo(task_sum.getRange(start, 19, 140,1), {contentsOnly:true});
      sheet_n.getRange("F3:G140").copyTo(task_sum.getRange(start, 20, 140,2), {contentsOnly:true});
      sheet_n.getRange("H3:I140").copyTo(task_sum.getRange(start, 22, 140,2), {contentsOnly:true});
      var array = task_sum.getRange("J3:J1000").getValues();
      var start = array.filter(String).length + 3; // removes the empty values and counts the length 
      task_sum.getRange(start + 3,8,140,1).setValue("");    
      
       //** Task Summary Complete
      var array = task_sum.getRange("Z3:Z1000").getValues();
      var start = array.filter(String).length + 3; // removes the empty values and counts the length
      var criteria = SpreadsheetApp.newFilterCriteria()
      .setHiddenValues(['Upcoming','Overdue', 'Scheduled', 'Ongoing', 'Not Scheduled']).build();    
      sheet_n.getFilter().setColumnFilterCriteria(5, criteria);
      //Populate Cmplete Columns
      sheet_n.getRange("B3:B140").copyTo(task_sum.getRange(start, 26, 140,1), {contentsOnly:true});
      sheet_n.getRange("D3:D140").copyTo(task_sum.getRange(start, 27, 140,1), {contentsOnly:true});
      sheet_n.getRange("F3:G140").copyTo(task_sum.getRange(start, 28, 140,2), {contentsOnly:true});
      sheet_n.getRange("H3:I140").copyTo(task_sum.getRange(start, 30, 140,2), {contentsOnly:true});
      var array = task_sum.getRange("Z3:Z1000").getValues();
      var start = array.filter(String).length + 3; // removes the empty values and counts the length 
      task_sum.getRange(start + 3,8,140,1).setValue("");    
           
       //** Task Summary Scheduled
      var array = task_sum.getRange("AH3:AH1000").getValues();
      var start = array.filter(String).length + 3; // removes the empty values and counts the length
      var criteria = SpreadsheetApp.newFilterCriteria()
      .setHiddenValues(['Upcoming','Overdue', 'Complete', 'Ongoing', 'Not Scheduled']).build();    
      sheet_n.getFilter().setColumnFilterCriteria(5, criteria);
      //Populate Scheduled Columns
      sheet_n.getRange("B3:B140").copyTo(task_sum.getRange(start, 34, 140,1), {contentsOnly:true});
      sheet_n.getRange("D3:D140").copyTo(task_sum.getRange(start, 35, 140,1), {contentsOnly:true});
      sheet_n.getRange("F3:G140").copyTo(task_sum.getRange(start, 36, 140,2), {contentsOnly:true});
      sheet_n.getRange("H3:I140").copyTo(task_sum.getRange(start, 38, 140,2), {contentsOnly:true});
      var array = task_sum.getRange("AH3:AH1000").getValues();
      var start = array.filter(String).length + 3; // removes the empty values and counts the length 
      task_sum.getRange(start + 3,8,140,1).setValue("");    
         
      
       //** Task Summary Not Scheduled
      var array = task_sum.getRange("AR3:AR1000").getValues();
      var start = array.filter(String).length + 3; // removes the empty values and counts the length
      var criteria = SpreadsheetApp.newFilterCriteria()
      .setHiddenValues(['Upcoming','Overdue', 'Complete', 'Ongoing', 'Not Scheduled']).build();    
      sheet_n.getFilter().setColumnFilterCriteria(5, criteria);     
      //Populate Not Schedulled columns
      sheet_n.getRange("B3:B140").copyTo(task_sum.getRange(start, 42, 140,1), {contentsOnly:true});
      sheet_n.getRange("D3:D140").copyTo(task_sum.getRange(start, 43, 140,1), {contentsOnly:true});
      sheet_n.getRange("F3:G140").copyTo(task_sum.getRange(start, 44, 140,2), {contentsOnly:true});
      sheet_n.getRange("H3:I140").copyTo(task_sum.getRange(start, 46, 140,2), {contentsOnly:true});
      task_sum.getRange("B1").setFormula('=COUNTIF(AP3:AP1000,"<>")');  
      task_sum.getRange(start,48,140,1).setValue(sht_name);
      var array = task_sum.getRange("AP3:AP1000").getValues();
      var start = array.filter(String).length + 3; // removes the empty values and counts the length 
      task_sum.getRange(start + 3,8,140,1).setValue("");    
         
      
      //************ GANTT CHART **************\
      //******* POPULATING THE GANTT CHART*****\\
      gantt.getRange("A2:G2229").setValue("");
  
      // Count the number of non-empty rows in the B column
      var array = gantt.getRange("B1:B2231").getValues();
      var start = array.filter(String).length; // removes the empty values and counts the length
      // Populate values
      sheet_n.getRange('B:B').createFilter();
      var criteria = SpreadsheetApp.newFilterCriteria()
      .setHiddenValues(['Planning', 'IRB', 'Project Close']).build();    
      sheet_n.getFilter().setColumnFilterCriteria(2, criteria)  
      sheet_n.getRange("B3:B140").copyTo(gantt.getRange(start +1 , 2, 140,1), {contentsOnly:true});
      sheet_n.getRange("D3:E140").copyTo(gantt.getRange(start +1, 3, 140,2), {contentsOnly:true});
      sheet_n.getRange("J3:J140").copyTo(gantt.getRange(start +1, 5, 140,1), {contentsOnly:true});
      sheet_n.getRange("F3:G140").copyTo(gantt.getRange(start +1, 6, 140,2), {contentsOnly:true});
      sheet_n.getFilter().remove();      
      gantt.getRange(start +1,1,140,1).setValue(sht_name);
      var array = gantt.getRange("B1:B2231").getValues();
      var start = array.filter(String).length; // removes the empty values and counts the length
      gantt.getRange(start+1,1,140,1).setValue("");
 
    }    
    
    
  }
  

}

