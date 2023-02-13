/************************************************************
Name:          AddChecklist
Purpose:       Triggers the html form that creates a new checklist
Author :       Mathew Bidinlib
email  :       mbidinlib@poverty-action.org
Date created:  Jan 20 2020
Date Modified: March 10 2020
Copyright    : Innovations for Poverty action @2020
***********************************************************/


function addchecklist(){
  function onOpen() {
    var ui = SpreadsheetApp.getUi();  
    ui.createMenu('Form').addItem('add Item', 'addItem').addToUi();
  }
  var html = HtmlService.createHtmlOutputFromFile('form');
  SpreadsheetApp.getUi().showModalDialog(html, 'Add New Checklist');
}


function addNewItem(form_data){
   // Define sheet names as variables
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName("master");
  var gantt = ss.getSheetByName("gantt");
  var startpage = ss.getSheetByName("Getting Started");
  var task_summary = ss.getSheetByName("task_summary");
  var empty_template = ss.getSheetByName("empty_template");
  var create_template = ss.getSheetByName("create_template")
  var type = form_data.checklist_type
  var sht = form_data.sheet_name
  
  Logger.log(type)
  Logger.log(sht)
  Logger.log(sht + "_" + type)
  
  
  // Flag error
  if(type == null){
     var action = Browser.msgBox("Checklist Error!","Checklist not selected", Browser.Buttons.OK);
       // Trigger html form again
     var html = HtmlService.createHtmlOutputFromFile('form');
     SpreadsheetApp.getUi().showModalDialog(html, 'Add New Checklist');
    
   }
  
  if(sht == ""){
    var action = Browser.msgBox("Entry Error!","Sheet name not entered", Browser.Buttons.OK);
      // Trigger html form again
     var html = HtmlService.createHtmlOutputFromFile('form');
     SpreadsheetApp.getUi().showModalDialog(html, 'Add New Checklist');
    
  }   
  
  // Create Checklist for Survey 
    if(type == "survey" && sht != ""){
       
      ///Create sheet and copy from template
      var NewSheet = ss.getSheetByName(sht + "_survey");       
      if (NewSheet != null) {
        var action = Browser.msgBox("Sheet Delete Alert!","The sheet you want to create already exists. Do you want to continue to replace it?", Browser.Buttons.OK_CANCEL);
        ss.deleteSheet(NewSheet);
      }
      
      NewSheet = ss.insertSheet();
      NewSheet.setName(sht + "_survey");
      var survey_sheet = ss.getSheetByName(sht + "_survey");
      create_template.getRange("A1:M140").copyTo(survey_sheet.getRange("A1:M140"))
      
       ss.moveActiveSheet(3); 
       ss.setActiveSheet(master);  
       var country = startpage.getRange("F66").getValue();
       var grds_input = ss.getSheetByName("grds_input");
      
       //*****Populating Created Implementation sheet****\\
    
      grds_input.getRange("A2:A200").copyTo(survey_sheet.getRange("C2:C200"),{contentsOnly:true});
      grds_input.getRange("C2:C200").copyTo(survey_sheet.getRange("B2:B200"),{contentsOnly:true});
      grds_input.getRange("F2:F200").copyTo(survey_sheet.getRange("F2:F200"),{contentsOnly:true});
      var num_rows = grds_input.getRange("T2").getValue();
      
        //Removing activities that dont belong here
      survey_sheet.getRange('F:F').createFilter();
      var criteria = SpreadsheetApp.newFilterCriteria().setHiddenValues(['template-General' ,'template-'+country]).build();
      survey_sheet.getFilter().setColumnFilterCriteria(6, criteria);
      survey_sheet.deleteRows(2,num_rows);
      survey_sheet.getRange('A2').clearDataValidations();
      survey_sheet.getRange('A13').clearDataValidations();
      survey_sheet.getFilter().remove();
      survey_sheet.getRange("F2:F140").setValue("");
             
    }
    
  
       //*** for implementation****//
    if(type == "impl" && sht != ""){
       
      ///Create sheet and copy from template
      var NewSheet = ss.getSheetByName(sht + "_impl");       
      if (NewSheet != null) {
        var action = Browser.msgBox("Sheet Delete Alert!","The sheet you want to create already exists. Do you want to continue to replace it?", Browser.Buttons.OK_CANCEL);
        ss.deleteSheet(NewSheet);
      }
      
      NewSheet = ss.insertSheet();
      NewSheet.setName(sht + "_impl");
      var imp_sheet = ss.getSheetByName(sht + "_impl");
      create_template.getRange("A1:M140").copyTo(imp_sheet.getRange("A1:M140"))
      
      ss.moveActiveSheet(3); 
      ss.setActiveSheet(master);
           
      var country = startpage.getRange("F66").getValue();
      var grds_input = ss.getSheetByName("grds_input");
      
  
      //*****Populating Created Implementation sheet****\\
    
      grds_input.getRange("A2:A200").copyTo(imp_sheet.getRange("C2:C200"),{contentsOnly:true});
      grds_input.getRange("C2:C200").copyTo(imp_sheet.getRange("B2:B200"),{contentsOnly:true});
      grds_input.getRange("F2:F200").copyTo(imp_sheet.getRange("F2:F200"),{contentsOnly:true});
      var num_rows = grds_input.getRange("T2").getValue();
      
            
         //Removing activities that dont belong here
      imp_sheet.getRange('F:F').createFilter();
      var criteria = SpreadsheetApp.newFilterCriteria().setHiddenValues(['implement-General' ,'implement-'+country]).build();
      imp_sheet.getFilter().setColumnFilterCriteria(6, criteria);
      imp_sheet.deleteRows(2,num_rows);
      
      imp_sheet.getRange('A2').clearDataValidations();
      imp_sheet.getRange('A13').clearDataValidations();
      imp_sheet.getFilter().remove();
      imp_sheet.getRange("F2:F140").setValue("");
      
  
    }
  
  //form empty tasks
    if(type == "empty" && sht != ""){
      
      ///Create sheet and copy from template
      var NewSheet = ss.getSheetByName(sht + "_task");       
      if (NewSheet != null) {
        var action = Browser.msgBox("Sheet Delete Alert!","The sheet you want to create already exists. Do you want to continue to replace it?", Browser.Buttons.OK_CANCEL);
        ss.deleteSheet(NewSheet);
      }
      
      NewSheet = ss.insertSheet();
      NewSheet.setName(sht + "_task");
      var task_sheet = ss.getSheetByName(sht + "_task");
      empty_template.getRange("A1:M140").copyTo(task_sheet.getRange("A1:M140"))
      
       ss.moveActiveSheet(3); 
       ss.setActiveSheet(master);
    }
    
  
  }
  
