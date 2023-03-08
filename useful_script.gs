/***********************************************************
Create Weekly subfolders for the curent Week. and Get the ID
This will skip folder creation if it already exisits
************************************************************/

const OUTPUT_FOLDER_URL = "https://drive.google.com/drive/u/0/folders/1xQIMgrW0ZaVw5iI4gqRENk1gdhwCXY4d"

var WeekFolderID = ''
var current_week = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("4 Progress Report").getRange("D2").getValue();
function createFolderSubfolder(folderURL, folderName) {
  folderID = folderURL.replace(/^.+\//, '')
  var pfolder = DriveApp.getFolderById(folderID);
  var wfolder = pfolder.getFoldersByName(folderName);
  if(wfolder.hasNext()) {
      WeekFolderID = wfolder.next().getId()
   } 
  else {
      var wfolder = pfolder.createFolder(folderName);
      WeekFolderID = wfolder.getId()
  }
}

createFolderSubfolder(OUTPUT_FOLDER_URL, current_week)




//https://developers.google.com/apps-script/samples/automations/generate-pdfs


// TODO: To test this solution, set EMAIL_OVERRIDE to true and set EMAIL_ADDRESS_OVERRIDE to your email address.
const EMAIL_OVERRIDE = true;
const EMAIL_ADDRESS_OVERRIDE = 'etsinigo@poverty-action.org';

// Application constants
const APP_TITLE = 'IPAG E-invoicing System';
const OUTPUT_FOLDER_NAME = "Consultants Invoices";
const OUTPUT_FOLDER_URL = "https://drive.google.com/drive/u/0/folders/1xQIMgrW0ZaVw5iI4gqRENk1gdhwCXY4d"

// Sheet name constants. Update if you change the names of the sheets.
const CONSULTANTS_SHEET_NAME = '1a Consultants';
const PRODUCTS_SHEET_NAME = '1b Fees Structure';
const TRANSACTIONS_SHEET_NAME = '5 Transactions';
const INVOICES_SHEET_NAME = '6 Invoices & Emails';
const INVOICE_TEMPLATE_SHEET_NAME = '6a Invoice Template';

//********************************************************************************************************
//Iterates through the worksheet data populating the template sheet with consultant data, then saves each instance as a PDF document. Called by user via custom menu item.
 
function processDocuments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const consultantsSheet = ss.getSheetByName(CONSULTANTS_SHEET_NAME);
  const productsSheet = ss.getSheetByName(PRODUCTS_SHEET_NAME);
  const transactionsSheet = ss.getSheetByName(TRANSACTIONS_SHEET_NAME);
  const invoicesSheet = ss.getSheetByName(INVOICES_SHEET_NAME);
  const invoiceTemplateSheet = ss.getSheetByName(INVOICE_TEMPLATE_SHEET_NAME);

  // Gets data from the storage sheets as objects.
  const consultants = dataRangeToObject(consultantsSheet);
  const products = dataRangeToObject(productsSheet);
  const transactions = dataRangeToObject(transactionsSheet);

  ss.toast('Creating Invoices', APP_TITLE, 1);
  const invoices = [];

  // Iterates for each consultant calling createInvoiceForCustomer routine.
  consultants.forEach(function (consultant) {
    ss.toast(`Creating Invoice for ${consultant.consultant_name}`, APP_TITLE, 1);
    let invoice = createInvoiceForConsultant(
      consultant, products, transactions, invoiceTemplateSheet, ss.getId());
    invoices.push(invoice);
  });
  // Writes invoices data to the sheet.
  invoicesSheet.getRange(2, 1, invoices.length, invoices[0].length).setValues(invoices);


  //Send emails
  sendEmails();
}

/***************************************************************************************************************
 * Processes each consultant instance with passed in data parameters.
 * @param {object} consultant - Object for the consultant
 * @param {object} products - Object for all the products
 * @param {object} transactions - Object for all the transactions
 * @param {object} invoiceTemplateSheet - Object for the invoice template sheet
 * @param {string} ssId - Google Sheet ID     
 * Return {array} of instance consultant invoice data
 */
function createInvoiceForConsultant(consultant, products, transactions, templateSheet, ssId) {
  let consultantTransactions = transactions.filter(function (transaction) {
    return transaction.ipa_id == consultant.ipa_id;
  });

  // Clears existing data from the template.
  clearTemplateSheet();

  // Calculate required values for the invoice
  let lineItems = [];
  let totalAmount = 0;
  let totalOutput = 0;
  let netAmount = 0;
  let taxAmount = 0;

  consultantTransactions.forEach(function (lineItem) {
    let lineItemProduct = products.filter(function (product) {
      return product.position == lineItem.position;
    })[0];
    const qty = parseInt(lineItem.outputs);
    const price = parseFloat(lineItemProduct.price).toFixed(2);
    const amount = parseFloat(qty * price).toFixed(2);
    const taxPay = parseFloat(amount * (7.5/100)).toFixed(2);
    const netAmt = parseFloat(amount - taxPay).toFixed(2);

    lineItems.push([lineItemProduct.position, lineItemProduct.milestones, '', '', qty, price, amount]);
    totalOutput += parseFloat(qty);
    totalAmount += parseFloat(amount);
    taxAmount += parseFloat(taxPay);
    netAmount += parseFloat(netAmt);
  
  });

  // Generates a random invoice number. You can replace with your own document ID method.
  const invoiceNumber = Math.floor(100000 + Math.random() * 900000);

  // Calulates dates.
  const todaysDate = new Date().toDateString()
  
  // Sets values in the template.
  const invoice_week = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("4 Progress Report").getRange(2,4).getValue();
  const projectName = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("4 Progress Report").getRange(2,7).getValue();
  const billPeriod = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("4 Progress Report").getRange(5,33).getValue();
  const data_phase = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("4 Progress Report").getRange(3,29).getValue();

  templateSheet.getRange('C2').setValue(consultant.ipa_id)
  templateSheet.getRange('F2').setValue(consultant.consultant_name) 

  templateSheet.getRange('C3').setValue(consultant.email)
  templateSheet.getRange('F3').setValue(consultant.phone)
  templateSheet.getRange('H3').setValue(consultant.position)  

  templateSheet.getRange('C9').setValue(invoiceNumber)
  templateSheet.getRange('E9').setValue(billPeriod)
  templateSheet.getRange('G9').setValue(invoice_week)
  templateSheet.getRange('I9').setValue(todaysDate)

  templateSheet.getRange('C7').setValue(projectName)
  templateSheet.getRange('H7').setValue(data_phase) 

  templateSheet.getRange('I31').setValue(consultant.initials)
  templateSheet.getRange('C31').setValue(consultant.consultant_name)

  templateSheet.getRange(12, 3, lineItems.length, 7).setValues(lineItems); //This command prefill the actual invoice data for quantity and outputs for each task and description.

 
  // Cleans up and creates PDF.
  SpreadsheetApp.flush();
  Utilities.sleep(500); // Using to offset any potential latency in creating .pdf
  const pdf = createPDF(ssId, templateSheet, `${consultant.ipa_id}: ${consultant.consultant_name} - Invoice No. ${invoiceNumber}`);
  return [consultant.ipa_id, consultant.consultant_name, consultant.position, invoiceNumber, todaysDate, consultant.email, totalOutput, '', '', totalAmount, taxAmount, netAmount, pdf.getUrl(), 'No'];

}

/*************************************************************************************************************************
* Resets the template sheet by clearing out consultant data.
* You use this to prepare for the next iteration or to view blank the template for design.
* Called by createInvoiceForCustomer() or by the user via custom menu item.
*/
function clearTemplateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = ss.getSheetByName(INVOICE_TEMPLATE_SHEET_NAME);
  
  // Clears existing data from the template.
  const rngClear = templateSheet.getRangeList(['C2:D3', 'F2', 'F3', 'H3', 'C7', 'H7', 'C9','E9', 'G9', 'I9', 'C12:I24', 'I31', 'C31']).getRanges() 
  rngClear.forEach(function (cell) {
    cell.clearContent();
  });
}


/***********************************************************
Create Weekly subfolders for the curent Week. and Get the ID
This will skip folder creation if it already exisits
************************************************************/

var WeekFolderID = ''
var current_week = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("4 Progress Report").getRange("D2").getValue();
function createFolderSubfolder(folderURL, folderName) {
  folderID = folderURL.replace(/^.+\//, '')
  var pfolder = DriveApp.getFolderById(folderID);
  var wfolder = pfolder.getFoldersByName(folderName);
  if(wfolder.hasNext()) {
      WeekFolderID = wfolder.next().getId()
   } 
  else {
      var wfolder = pfolder.createFolder(folderName);
      WeekFolderID = wfolder.getId()
  }
}

createFolderSubfolder(OUTPUT_FOLDER_URL, current_week)


/****************************************************************************************************************************
 * Creates a PDF for the consultant given sheet.
 * @param {string} ssId - Id of the Google Spreadsheet
 * @param {object} sheet - Sheet to be converted as PDF
 * @param {string} pdfName - File name of the PDF being created
 * @return {file object} PDF file as a blob
 */
function createPDF(ssId, sheet, pdfName) {
  const fr = 0, fc = 0, lc = 10, lr = 33;
  const url = "https://docs.google.com/spreadsheets/d/" + ssId + "/export" +
    "?format=pdf&" +
    "size=A4&" +
    "fzr=true&" +
    "portrait=false&" +
    "fitw=true&" +
    "gridlines=true&" +
    "printtitle=false&" +
    "top_margin=0.1&" +
    "bottom_margin=0.1&" +
    "left_margin=0.1&" +
    "right_margin=0.1&" +
    "sheetnames=false&" +
    "pagenum=true&" +
    "attachment=true&" +
    "gid=" + sheet.getSheetId() + '&' +
    "r1=" + fr + "&c1=" + fc + "&r2=" + lr + "&c2=" + lc;

  const params = { method: "GET", headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } };
  const blob = UrlFetchApp.fetch(url, params).getBlob().setName(pdfName + '.pdf');

  // Gets the folder in Drive where the PDFs are stored.
  //const folder = getFolderByName_(OUTPUT_FOLDER_NAME);
  const folder = DriveApp.getFolderById(WeekFolderID)

  
  /************************
     Replace file if exists
     Turn on if neccesary
  **************************/
  /*
  currfile = folder.getFilesByName(pdfName + '.pdf');
  if (currfile.hasNext()) {//If there is another element in the iterator
    currfileID = currfile.next().getId();
    rtrnFromDLET = Drive.Files.remove(currfileID)
  };
  */



  const pdfFile = folder.createFile(blob);
  return pdfFile;
}


/************************************************************************************************************************************
 * Sends emails to field consultants with PDF as an attachment
 * **********************************************************************************************************************************
 * Checks/Sets 'Email Sent' column to 'Yes' to avoid resending.
 * Called by user via custom menu item.
 */
function sendEmails() {

  var ui = SpreadsheetApp.getUi() 
  var presponse = ui.alert("Send Emails?", 'This will send emails to field consultants with thier validated invoices',ui.ButtonSet.YES_NO);
  if (presponse == ui.Button.NO) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invoicesSheet = ss.getSheetByName(INVOICES_SHEET_NAME);

  const invoicesData = invoicesSheet.getRange(1, 1, invoicesSheet.getLastRow(), invoicesSheet.getLastColumn()).getValues();
  const keysI = invoicesData.splice(0, 1)[0];
  const invoices = getObjects(invoicesData, createObjectKeys(keysI));
  const templateText = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("1 Settings").getRange(1700,32).getValue();
  const billingPeriod = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("4 Progress Report").getRange(5,33).getValue();

  ss.toast('Emailing Invoices', APP_TITLE, 1);
  invoices.forEach(function (invoice, index) {
    if (invoice.email_sent != 'Yes') {
      ss.toast(`Emailing Invoice for ${invoice.consultant_name}`, APP_TITLE, 1);
      
      const projectName = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("4 Progress Report").getRange(2,7).getValue();
      const invoiceweek = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("4 Progress Report").getRange(2,4).getValue();

      const name = invoice.consultant_name
      const totAmnt = invoice.total_amount
      const totOutput = invoice.total_output
      const invoiceN = invoice.invoice_no
      const role = invoice.position
      const tax = invoice.tax
      const netTotal = invoice.amount_due

      const fileId = invoice.invoice_link.match(/[-\w]{25,}(?!.*[-\w]{25,})/)
      const attachment = DriveApp.getFileById(fileId);
      const EMAIL_SUBJECT = 'IPAG: ' + 'Payment for Invoice #' + invoiceN + ' for Services Rendered as ' + role 

      var messageBody = templateText.replace("{name}",name).replace("{role}",role).replace("{period}",billingPeriod).replace("{weekn}",invoiceweek).replace("{bill}",invoiceN).replace("{title}",projectName).replace("{amt}",netTotal);

      let recipient = invoice.email;
      if (EMAIL_OVERRIDE) {
        recipient = EMAIL_ADDRESS_OVERRIDE
      }
      
        MailApp.sendEmail(recipient, EMAIL_SUBJECT,"", {
        htmlBody: messageBody.replace(/\n/g,'<br>'), 
        attachments: [attachment.getAs(MimeType.PDF)],
        name: APP_TITLE
        });

      invoicesSheet.getRange(index + 2, 14).setValue('Yes');
    }
  });
}





/*******************************************************************************************************
 * Helper function that turns sheet data range into an object. 
 * @param {SpreadsheetApp.Sheet} sheet - Sheet to process
 * Return {object} of a sheet's datarange as an object 
 */
function dataRangeToObject(sheet) {
  const dataRange = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  const keys = dataRange.splice(0, 1)[0];
  return getObjects(dataRange, createObjectKeys(keys));
}

/********************************************************************************************************
 * Utility function for mapping sheet data to objects.
 */
function getObjects(data, keys) {
  let objects = [];
  for (let i = 0; i < data.length; ++i) {
    let object = {};
    let hasData = false;
    for (let j = 0; j < data[i].length; ++j) {
      let cellData = data[i][j];
      if (isCellEmpty(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}
// Creates object keys for column headers.
function createObjectKeys(keys) {
  return keys.map(function (key) {
    return key.replace(/\W+/g, '_').toLowerCase();
  });
}
// Returns true if the cell where cellData was read from is empty.
function isCellEmpty(cellData) {
  return typeof (cellData) == "string" && cellData == "";
}












//Edward Tsinigo -- Based on initial script from 
//https://developers.google.com/apps-script/samples/automations/generate-pdfs


// TODO: To test this solution, set EMAIL_OVERRIDE to true and set EMAIL_ADDRESS_OVERRIDE to your email address.
const EMAIL_OVERRIDE = true;
const EMAIL_ADDRESS_OVERRIDE = 'etsinigo@poverty-action.org';

// Application constants
const APP_TITLE = 'IPAG E-invoicing System';
const OUTPUT_FOLDER_NAME = "Consultants Invoices";
const OUTPUT_FOLDER_URL = "https://drive.google.com/drive/u/0/folders/1xQIMgrW0ZaVw5iI4gqRENk1gdhwCXY4d"

// Sheet name constants. Update if you change the names of the sheets.
const CONSULTANTS_SHEET_NAME = '1a Consultants';
const PRODUCTS_SHEET_NAME = '1b Fees Structure';
const TRANSACTIONS_SHEET_NAME = '5 Transactions';
const INVOICES_SHEET_NAME = '6 Invoices & Emails';
const INVOICE_TEMPLATE_SHEET_NAME = '6a Invoice Template';

//********************************************************************************************************
//Iterates through the worksheet data populating the template sheet with consultant data, then saves each instance as a PDF document. Called by user via custom menu item.
 
function processDocuments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const consultantsSheet = ss.getSheetByName(CONSULTANTS_SHEET_NAME);
  const productsSheet = ss.getSheetByName(PRODUCTS_SHEET_NAME);
  const transactionsSheet = ss.getSheetByName(TRANSACTIONS_SHEET_NAME);
  const invoicesSheet = ss.getSheetByName(INVOICES_SHEET_NAME);
  const invoiceTemplateSheet = ss.getSheetByName(INVOICE_TEMPLATE_SHEET_NAME);

  // Gets data from the storage sheets as objects.
  const consultants = dataRangeToObject(consultantsSheet);
  const products = dataRangeToObject(productsSheet);
  const transactions = dataRangeToObject(transactionsSheet);

  ss.toast('Creating Invoices', APP_TITLE, 1);
  const invoices = [];

  // Iterates for each consultant calling createInvoiceForCustomer routine.
  consultants.forEach(function (consultant) {
    ss.toast(`Creating Invoice for ${consultant.consultant_name}`, APP_TITLE, 1);
    let invoice = createInvoiceForConsultant(
      consultant, products, transactions, invoiceTemplateSheet, ss.getId());
    invoices.push(invoice);
  });
  // Writes invoices data to the sheet.
  invoicesSheet.getRange(2, 1, invoices.length, invoices[0].length).setValues(invoices);


  //Send emails
  sendEmails();
}

/***************************************************************************************************************
 * Processes each consultant instance with passed in data parameters.
 * @param {object} consultant - Object for the consultant
 * @param {object} products - Object for all the products
 * @param {object} transactions - Object for all the transactions
 * @param {object} invoiceTemplateSheet - Object for the invoice template sheet
 * @param {string} ssId - Google Sheet ID     
 * Return {array} of instance consultant invoice data
 */
function createInvoiceForConsultant(consultant, products, transactions, templateSheet, ssId) {
  let consultantTransactions = transactions.filter(function (transaction) {
    return transaction.ipa_id == consultant.ipa_id;
  });

  // Clears existing data from the template.
  clearTemplateSheet();

  // Calculate required values for the invoice
  let lineItems = [];
  let totalAmount = 0;
  let totalOutput = 0;
  let netAmount = 0;
  let taxAmount = 0;

  consultantTransactions.forEach(function (lineItem) {
    let lineItemProduct = products.filter(function (product) {
      return product.position == lineItem.position;
    })[0];
    const qty = parseInt(lineItem.outputs);
    const price = parseFloat(lineItemProduct.price).toFixed(2);
    const amount = parseFloat(qty * price).toFixed(2);
    const taxPay = parseFloat(amount * (7.5/100)).toFixed(2);
    const netAmt = parseFloat(amount - taxPay).toFixed(2);

    lineItems.push([lineItemProduct.position, lineItemProduct.milestones, '', '', qty, price, amount]);
    totalOutput += parseFloat(qty);
    totalAmount += parseFloat(amount);
    taxAmount += parseFloat(taxPay);
    netAmount += parseFloat(netAmt);
  
  });

  // Generates a random invoice number. You can replace with your own document ID method.
  const invoiceNumber = Math.floor(100000 + Math.random() * 900000);

  // Calulates dates.
  const todaysDate = new Date().toDateString()
  
  // Sets values in the template.
  const invoice_week = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("4 Progress Report").getRange(2,4).getValue();
  const projectName = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("4 Progress Report").getRange(2,7).getValue();
  const billPeriod = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("4 Progress Report").getRange(5,33).getValue();
  const data_phase = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("4 Progress Report").getRange(3,29).getValue();

  templateSheet.getRange('C2').setValue(consultant.ipa_id)
  templateSheet.getRange('F2').setValue(consultant.consultant_name) 

  templateSheet.getRange('C3').setValue(consultant.email)
  templateSheet.getRange('F3').setValue(consultant.phone)
  templateSheet.getRange('H3').setValue(consultant.position)  

  templateSheet.getRange('C9').setValue(invoiceNumber)
  templateSheet.getRange('E9').setValue(billPeriod)
  templateSheet.getRange('G9').setValue(invoice_week)
  templateSheet.getRange('I9').setValue(todaysDate)

  templateSheet.getRange('C7').setValue(projectName)
  templateSheet.getRange('H7').setValue(data_phase) 

  templateSheet.getRange('I31').setValue(consultant.initials)
  templateSheet.getRange('C31').setValue(consultant.consultant_name)

  templateSheet.getRange(12, 3, lineItems.length, 7).setValues(lineItems); //This command prefill the actual invoice data for quantity and outputs for each task and description.

 
  // Cleans up and creates PDF.
  SpreadsheetApp.flush();
  Utilities.sleep(500); // Using to offset any potential latency in creating .pdf
  const pdf = createPDF(ssId, templateSheet, `${consultant.ipa_id}: ${consultant.consultant_name} - Invoice No. ${invoiceNumber}`);
  return [consultant.ipa_id, consultant.consultant_name, consultant.position, invoiceNumber, todaysDate, consultant.email, totalOutput, '', '', totalAmount, taxAmount, netAmount, pdf.getUrl(), 'No'];

}

/*************************************************************************************************************************
* Resets the template sheet by clearing out consultant data.
* You use this to prepare for the next iteration or to view blank the template for design.
* Called by createInvoiceForCustomer() or by the user via custom menu item.
*/
function clearTemplateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = ss.getSheetByName(INVOICE_TEMPLATE_SHEET_NAME);
  
  // Clears existing data from the template.
  const rngClear = templateSheet.getRangeList(['C2:D3', 'F2', 'F3', 'H3', 'C7', 'H7', 'C9','E9', 'G9', 'I9', 'C12:I24', 'I31', 'C31']).getRanges() 
  rngClear.forEach(function (cell) {
    cell.clearContent();
  });
}


/***********************************************************
Create Weekly subfolders for the curent Week. and Get the ID
This will skip folder creation if it already exisits
************************************************************/

var WeekFolderID = ''
var current_week = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("4 Progress Report").getRange("D2").getValue();
function createFolderSubfolder(folderURL, folderName) {
  folderID = folderURL.replace(/^.+\//, '')
  var pfolder = DriveApp.getFolderById(folderID);
  var wfolder = pfolder.getFoldersByName(folderName);
  if(wfolder.hasNext()) {
      WeekFolderID = wfolder.next().getId()
   } 
  else {
      var wfolder = pfolder.createFolder(folderName);
      WeekFolderID = wfolder.getId()
  }
}

createFolderSubfolder(OUTPUT_FOLDER_URL, current_week)


/****************************************************************************************************************************
 * Creates a PDF for the consultant given sheet.
 * @param {string} ssId - Id of the Google Spreadsheet
 * @param {object} sheet - Sheet to be converted as PDF
 * @param {string} pdfName - File name of the PDF being created
 * @return {file object} PDF file as a blob
 */
function createPDF(ssId, sheet, pdfName) {
  const fr = 0, fc = 0, lc = 10, lr = 33;
  const url = "https://docs.google.com/spreadsheets/d/" + ssId + "/export" +
    "?format=pdf&" +
    "size=A4&" +
    "fzr=true&" +
    "portrait=false&" +
    "fitw=true&" +
    "gridlines=true&" +
    "printtitle=false&" +
    "top_margin=0.1&" +
    "bottom_margin=0.1&" +
    "left_margin=0.1&" +
    "right_margin=0.1&" +
    "sheetnames=false&" +
    "pagenum=true&" +
    "attachment=true&" +
    "gid=" + sheet.getSheetId() + '&' +
    "r1=" + fr + "&c1=" + fc + "&r2=" + lr + "&c2=" + lc;

  const params = { method: "GET", headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } };
  const blob = UrlFetchApp.fetch(url, params).getBlob().setName(pdfName + '.pdf');

  // Gets the folder in Drive where the PDFs are stored.
  //const folder = getFolderByName_(OUTPUT_FOLDER_NAME);
  const folder = DriveApp.getFolderById(WeekFolderID)

  
  /************************
     Replace file if exists
     Turn on if neccesary
  **************************/
  /*
  currfile = folder.getFilesByName(pdfName + '.pdf');
  if (currfile.hasNext()) {//If there is another element in the iterator
    currfileID = currfile.next().getId();
    rtrnFromDLET = Drive.Files.remove(currfileID)
  };
  */



  const pdfFile = folder.createFile(blob);
  return pdfFile;
}


/************************************************************************************************************************************
 * Sends emails to field consultants with PDF as an attachment
 * **********************************************************************************************************************************
 * Checks/Sets 'Email Sent' column to 'Yes' to avoid resending.
 * Called by user via custom menu item.
 */
function sendEmails() {

  var ui = SpreadsheetApp.getUi() 
  var presponse = ui.alert("Send Emails?", 'This will send emails to field consultants with thier validated invoices',ui.ButtonSet.YES_NO);
  if (presponse == ui.Button.NO) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invoicesSheet = ss.getSheetByName(INVOICES_SHEET_NAME);

  const invoicesData = invoicesSheet.getRange(1, 1, invoicesSheet.getLastRow(), invoicesSheet.getLastColumn()).getValues();
  const keysI = invoicesData.splice(0, 1)[0];
  const invoices = getObjects(invoicesData, createObjectKeys(keysI));
  const templateText = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("1 Settings").getRange(1700,32).getValue();
  const billingPeriod = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("4 Progress Report").getRange(5,33).getValue();

  ss.toast('Emailing Invoices', APP_TITLE, 1);
  invoices.forEach(function (invoice, index) {
    if (invoice.email_sent != 'Yes') {
      ss.toast(`Emailing Invoice for ${invoice.consultant_name}`, APP_TITLE, 1);
      
      const projectName = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("4 Progress Report").getRange(2,7).getValue();
      const invoiceweek = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("4 Progress Report").getRange(2,4).getValue();

      const name = invoice.consultant_name
      const totAmnt = invoice.total_amount
      const totOutput = invoice.total_output
      const invoiceN = invoice.invoice_no
      const role = invoice.position
      const tax = invoice.tax
      const netTotal = invoice.amount_due

      const fileId = invoice.invoice_link.match(/[-\w]{25,}(?!.*[-\w]{25,})/)
      const attachment = DriveApp.getFileById(fileId);
      const EMAIL_SUBJECT = 'IPAG: ' + 'Payment for Invoice #' + invoiceN + ' for Services Rendered as ' + role 

      var messageBody = templateText.replace("{name}",name).replace("{role}",role).replace("{period}",billingPeriod).replace("{weekn}",invoiceweek).replace("{bill}",invoiceN).replace("{title}",projectName).replace("{amt}",netTotal);

      let recipient = invoice.email;
      if (EMAIL_OVERRIDE) {
        recipient = EMAIL_ADDRESS_OVERRIDE
      }
      
        MailApp.sendEmail(recipient, EMAIL_SUBJECT,"", {
        htmlBody: messageBody.replace(/\n/g,'<br>'), 
        attachments: [attachment.getAs(MimeType.PDF)],
        name: APP_TITLE
        });

      invoicesSheet.getRange(index + 2, 14).setValue('Yes');
    }
  });
}





/*******************************************************************************************************
 * Helper function that turns sheet data range into an object. 
 * @param {SpreadsheetApp.Sheet} sheet - Sheet to process
 * Return {object} of a sheet's datarange as an object 
 */
function dataRangeToObject(sheet) {
  const dataRange = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  const keys = dataRange.splice(0, 1)[0];
  return getObjects(dataRange, createObjectKeys(keys));
}

/********************************************************************************************************
 * Utility function for mapping sheet data to objects.
 */
function getObjects(data, keys) {
  let objects = [];
  for (let i = 0; i < data.length; ++i) {
    let object = {};
    let hasData = false;
    for (let j = 0; j < data[i].length; ++j) {
      let cellData = data[i][j];
      if (isCellEmpty(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}
// Creates object keys for column headers.
function createObjectKeys(keys) {
  return keys.map(function (key) {
    return key.replace(/\W+/g, '_').toLowerCase();
  });
}
// Returns true if the cell where cellData was read from is empty.
function isCellEmpty(cellData) {
  return typeof (cellData) == "string" && cellData == "";
}


