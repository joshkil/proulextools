/**
 * Proulex is a language learning school (https://www.proulex.com/). 
 * The developer of this script works for Proulex. The script provides 
 * Add On tools for Google Sheets to help teachers distribute test results
 * from tests administered through Google Forms. 
 * 
 * Author: kilpatrick.joshua@gmail.com
 */

/**
 * A special AppScript function that runs when the spreadsheet is first
 * opened or reloaded. onOpen() is used to add custom menu
 * items to the spreadsheet.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Proulex Tools')
    .addItem(
      'Create Test Result Handouts', 'processProulexTestFromResponses')
    .addToUi();
}

function processProulexTestFromResponses() {
  var spreadsheet = SpreadsheetApp.getActive();

  spreadsheet.toast("Your form responses are being processed. The process can take a minute. We'll let you know when the results are ready.", "Proulex Tools", 30);

  /**
   * Create a copy of the form responses table in a new sheet titled "Data"
   */
  {
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");
    if (sheet != null) {
      SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheet);
    }
  }
  spreadsheet.setActiveSheet(spreadsheet.getSheets()[0]);
  dataSheet = spreadsheet.duplicateActiveSheet();
  dataSheet.setName("Data");

  /**
   * Create a new sheet for holding the results URL's for individual students
   */
  {
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Student Files");
    if (sheet != null) {
      SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheet);
    }
  }
  studentFilesSheet = spreadsheet.insertSheet("Student Files");
  studentFilesSheet.setColumnWidths(1, 2, 300);
  studentFilesSheet.getRange('A1:B1').setBackground('#000000')
  .setFontColor('#ffffff')
  .setFontWeight('bold');
  studentFilesSheet.getRange(1, 1).setValue("Student Name");
  studentFilesSheet.getRange(1, 2).setValue("Student Results Doc");

  /**
   * Get a range that is the first two rows of the Google Form Results Table.
   * This is the header row with the questions in the form, and the first 
   * row of responses. 
   */
  var allrows = dataSheet.getDataRange();
  var tworows = allrows.offset(0,0,2);
  var numResponses = allrows.getHeight() - 1;
  Logger.log(numResponses);

  /**
   * Loop through the number of form responses included in the sheet. 
   * Process each form resopnse and write the PDF URL into the 
   * studentFilesSheet. 
   */
  for(i = 0; i < numResponses; i++){
    var studentName = tworows.getValues()[1][2];
    var studentFile = processOneStudent(spreadsheet, tworows);
    studentFilesSheet.getRange(2 + i, 1).setValue(studentName);
    studentFilesSheet.getRange(2 + i, 2).setValue(studentFile.getUrl());
    SpreadsheetApp.flush();
    dataSheet.deleteRows(2,1);
  }
  
  /**
   * Delete the dataSheet which was just a copy of the form response sheet 
   * used for processing. 
   */
  spreadsheet.deleteSheet(dataSheet);

  spreadsheet.toast("We finished processing your results. Look at the new \"Student Files\" sheet to find URLs you can share with your students.", "Proulex Tools", 30);
};

/**
 * Processes one student's form responses and save a PDF with those results.
 * @param {spreadsheet object} spreadsheet - the current spreadsheet
 * @param {range object} twoRows - a Range containing two rows. 
 * @return {file object} PDF file as a blob
 */
function processOneStudent(spreadsheet, tworows){

  var data = tworows.getValues();
  var studentName = data[1][2];

  /**
   * Check if a sheet already exsits with the Student's Name and delete it. 
   */
  {
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(studentName);
    if (sheet != null) {
      SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheet);
    }
  }

  /**
   * Create a new sheet to transpose and format the student's results. 
   */
  var studentSheet = spreadsheet.insertSheet(studentName);
  var cell = studentSheet.getRange('A1');
  var formula = '=TRANSPOSE(\'Data\'!'.concat(tworows.getA1Notation()).concat(')');
  cell.setFormula(formula);

  stDataRange = studentSheet.getDataRange();
  stDataRange.copyValuesToRange(stDataRange.getGridId(), 3, 3 + stDataRange.getWidth(), 1, stDataRange.getHeight());
  stDataRange.deleteCells(SpreadsheetApp.Dimension.COLUMNS);

  stDataRange = studentSheet.getDataRange();
  stDataRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  studentSheet.getRange('A3:B3').setBackground('#000000')
  .setFontColor('#ffffff')
  .setFontWeight('bold');

  studentSheet.setColumnWidths(1, 2, 436);
  studentSheet.getRange('B:B').setHorizontalAlignment('left');

  /**
   * Cleans up and creates PDF
   */
  SpreadsheetApp.flush();
  Utilities.sleep(500); // Using to offset any potential latency in creating .pdf

  studentFile = createPDF(spreadsheet, studentSheet);
  Logger.log(studentName + ": " + studentFile.getUrl());
  /**
   * Delete the student sheet where the student's reponses were
   * transposed to make the PDF. 
   */
  spreadsheet.deleteSheet(studentSheet);
  return studentFile;
};

/**
 * Returns a Google Drive folder in the same location 
 * in Drive where the spreadsheet is located. First, it checks if the folder
 * already exists and returns that folder. If the folder doesn't already
 * exist, the script creates a new one. 
 *
 * @param {string} folderName - Name of the Drive folder to create. 
 * @return {object} Google Drive Folder
 */
function getFolderByName_(folderName) {

  // Gets the Drive Folder of where the current spreadsheet is located.
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const parentFolder = DriveApp.getFileById(ssId).getParents().next();

  // Iterates the subfolders to check if the PDF folder already exists.
  const subFolders = parentFolder.getFolders();
  while (subFolders.hasNext()) {
    let folder = subFolders.next();

    // Returns the existing folder if found.
    if (folder.getName() === folderName) {
      return folder;
    }
  }
  // Creates a new folder if one does not already exist.
  return parentFolder.createFolder(folderName)
    .setDescription(`Created by Proulex Tools AddOn to store PDF output files`);
}

/**
 * Creates a PDF from the specified spreadsheet sheet. 
 * @param {object} ss - the Google Spreadsheet
 * @param {object} sheet - the Sheet to be exported as PDF
 * @return {file object} PDF file as a blob
 */
function createPDF(ss, sheet) {
  /**
   * These constants define the first row and column and the last row and column
   * to include in the pdf. 
   */
  const fr = 0, fc = 0;
  const lc = sheet.getDataRange().getWidth();
  const lr = sheet.getDataRange().getHeight();

  /**
   * This URL exports a given sheet as a PDF. 
   * This github was very helpful for understanding the URL parameters. 
   * https://gist.github.com/Spencer-Easton/78f9867a691e549c9c70
   */
  const url = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export" +
    "?format=pdf&" +
    "size=0&" +
    "fzr=true&" +
    "portrait=true&" +
    "fitw=true&" +
    "gridlines=true&" +
    "printtitle=false&" +
    "top_margin=0.5&" +
    "bottom_margin=0.25&" +
    "left_margin=0.5&" +
    "right_margin=0.5&" +
    "sheetnames=false&" +
    "pagenum=UNDEFINED&" +
    "attachment=true&" +
    "gid=" + sheet.getSheetId() + '&' +
    "r1=" + fr + "&c1=" + fc + "&r2=" + lr + "&c2=" + lc;

  /**
   * This code was taken from a sample script in Google Workplace Developers pages. 
   * https://developers.google.com/apps-script/samples/automations/generate-pdfs
   */

  const params = { method: "GET", headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } };
  const blob = UrlFetchApp.fetch(url, params).getBlob().setName(sheet.getName() + '.pdf');

  // Gets the folder in Drive where the PDF will be stored.
  const folder = getFolderByName_(ss.getName());

  const pdfFile = folder.createFile(blob);
  return pdfFile;
};
