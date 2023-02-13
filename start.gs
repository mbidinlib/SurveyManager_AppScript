/************************************************************
Name:          Start Using Myra
Purpose:       Sets up a project MyRA at the initial stage
Author :       Mathew Bidinlib
email  :       mbidinlib@poverty-action.org
Date created:  Jan 20 2020
Date Modified: March 10 2020
Copyright    : Innovations for Poverty action @2020
***********************************************************/

function start(){
  // Define sheet names as variables
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName("master");
  var gantt = ss.getSheetByName("gantt");
  var startpage = ss.getSheetByName("Getting Started");
  var task_summary = ss.getSheetByName("task_summary");
  var empty_template = ss.getSheetByName("empty_template");
  var create_template = ss.getSheetByName("create_template");
  var grds_input = ss.getSheetByName("grds_input");
  
 
   // Show and hide sheets
  master.showSheet();
  ss.setActiveSheet(master);
  task_summary.showSheet();
  gantt.showSheet();
  startpage.hideSheet() 
  
  //  Create project open and close sheets
  var p_open = ss.insertSheet();
  p_open .setName("project_open");
  var new_sheet = ss.getSheetByName("project_open");
  create_template.getRange("A1:M140").copyTo(new_sheet.getRange("A1:M140"));
  
  var p_close = ss.insertSheet();
  p_close .setName("project_close");
  var new_sheet = ss.getSheetByName("project_close");
  create_template.getRange("A1:M140").copyTo(new_sheet.getRange("A1:M140"));
  
  ss.setActiveSheet(master); // make master sheet active
    
  // Delete prefilled data if existis
  master.getRange("C3:C11").setValue("");
  master.getRange("E16:F65").setValue("");
  // Set CountryName
  var country = startpage.getRange("F66").getValue();
  master.getRange("C2").setValue(country);
  // add RQ and Finpro detials for Ghana in the master
  if(country == "Ghana"){
    
         master.getRange("E12").setValue("Ghana Research Quality Team");
         master.getRange("E13").setValue("Ghana Finance and Procurement Team");
         master.getRange("F12").setValue("mbidinlib@poverty-action.org");
         master.getRange("F13").setValue("finprocgh@poverty-action.org");
         master.getRange("E12:E13").setBackground("#d9e2f3")
         master.getRange("E12:F13").setBorder(true, true, true, true, true, true)
   }
     
  // Populate Project Open Sheets
  var project_open = ss.getSheetByName("project_open");
  var project_close = ss.getSheetByName("project_close");
  var range_gi = grds_input.getRange("A2:A200");
  var range_gi1 = grds_input.getRange("C2:C200");
  var range_gi2 = grds_input.getRange("F2:F200");
  
  //**** Populating project Open****\\ 
  range_gi.copyTo(project_open.getRange("C2:C200"));
  range_gi1.copyTo(project_open.getRange("B2:B200"));
  range_gi2.copyTo(project_open.getRange("F2:F200"), {contentsOnly:true});
  var num_rows = grds_input.getRange("T2").getValue();   
       
      //Keep only Project open rows
  project_open.getRange('F:F').createFilter();
  var criteria = SpreadsheetApp.newFilterCriteria().setHiddenValues(['project_open-General' ,'project_open-'+country]).build();
  project_open.getFilter().setColumnFilterCriteria(6, criteria);
  project_open.deleteRows(2,num_rows);
  project_open.getRange('A2').clearDataValidations();
  project_open.getRange('A13').clearDataValidations();
  project_open.getFilter().removeColumnFilterCriteria(6);
  project_open.getRange("F2:F140").setValue("");
  
 
  //*****Populating project close sheet****\\
  //***************************************\\ 
      //Copy all to sheet
  range_gi.copyTo(project_close.getRange("C2:C200"));
  range_gi1.copyTo(project_close.getRange("B2:B200"));
  range_gi2.copyTo(project_close.getRange("F2:F200"),{contentsOnly:true});
  var num_rows = grds_input.getRange("T2").getValue();
  
        //Keep only Project open rows
  project_close.getRange('F:F').createFilter();
  var criteria = SpreadsheetApp.newFilterCriteria().setHiddenValues(['project_close-General' ,'project_close-'+country]).build();
  project_close.getFilter().setColumnFilterCriteria(6, criteria);
  project_close.deleteRows(2,num_rows);
  
  project_close.getRange('A2').clearDataValidations();
  project_close.getRange('A13').clearDataValidations();
  project_close.getFilter().removeColumnFilterCriteria(6);
  project_close.getRange("F2:F140").setValue("");
  
}
