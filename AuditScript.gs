var auditSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var checklistSheet = auditSpreadsheet.getSheetByName("Checklist");
var detailsSheet = auditSpreadsheet.getSheetByName("Details");
var scopeSheet = auditSpreadsheet.getSheetByName("Scope");
var folderID = '1j24G3roDm8ySrkg4NvHL--__6ZNTm9Hk'; // set to ID of the folder ex. https://drive.google.com/drive/folders/<id of folder>

function onOpen() {
  setAuditTypes();
  createUI();
}

function createUI() {
  var ui = SpreadsheetApp.getUi();
  
  if (folderID !== '') {
    ui.createMenu(folderID)
      .addItem('Copy to Directory', 'copyToAuditDirectory');
  }
  ui.createMenu('Audit')
    .addItem('Create checklists from Scope', 'createChecklists')
//    .addItem('Copy to Directory', 'copyToAuditDirectory')
    .addToUi();
}

function createChecklists() {
  // make a copy of the checklist we can safely modify for bulk duplicating
  checklistSheet.copyTo(auditSpreadsheet).setName("Checklist Template");
  var checklistTemplateSheet = auditSpreadsheet.getSheetByName("Checklist Template");

  // add "Issue Tracker" "Status" and "Issues" columns to the checklist template
  checklistTemplateSheet.insertColumns(1, 3); // at position 1, insert 2 columns
  checklistTemplateSheet.getRange("A1").setValue("Issue Tracker");
  checklistTemplateSheet.getRange("B1").setValue("Status");
  checklistTemplateSheet.getRange("C1").setValue("Issues");
  
  // setting default values and data validation for Status
  var checklistTemplateLastRow = checklistTemplateSheet.getLastRow();
  var statusColumn = checklistTemplateSheet.getRange(2,2,checklistTemplateLastRow - 1);
  var statusColumnValidation = SpreadsheetApp.newDataValidation().requireValueInList(['Check Incomplete','Pass','Fail','Not Applicable','Notes / Other']).build();
  statusColumn.setValue('Check Incomplete');
  statusColumn.setDataValidation(statusColumnValidation);
  
  // sizing created columns
  checklistTemplateSheet.autoResizeColumn(2); // should be the "Status" column
  checklistTemplateSheet.setColumnWidth(3, 400); // should be the "Issues" column
  
  // filtering the checklist template by the chosen scope
  var auditType = detailsSheet.getRange('B1').getValue(); // future - set to be relative to B column when A column that has "Audit Type"

  var checklistTemplateSheetHeaders = checklistTemplateSheet.getRange("Checklist Template!1:1").getValues(); // first row of checklistTemplateSheet
  var auditTypeHeader = "Audit - " + auditType;
  
  var filterColumn = "";
  for (var i=0; i < checklistTemplateSheetHeaders[0].length; i++) {
    if (checklistTemplateSheetHeaders[0][i] == auditTypeHeader) {
      filterColumn = i+1;
    }
  }
  var filterColumnLetter = columnToLetter(filterColumn);
  var rangeString = "Checklist Template!" + filterColumnLetter + ":" + filterColumnLetter;  
  var auditFilter = checklistTemplateSheet.getRange(rangeString).createFilter();
  var auditFilterCriteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['FALSE'])
  .build()
  checklistTemplateSheet.getFilter().setColumnFilterCriteria(filterColumn, auditFilterCriteria);
  
   
  // creating a Checklist sheet for each row in the "Scope" sheet

  scopeRange = scopeSheet.getDataRange();
  scopeData = scopeRange.getValues();
  lastRow = scopeRange.getLastRow();
  
  for (var i=1; i < lastRow; i++) {
    var name = scopeData[i][0];
    // var url = scopeData[i][1]; // future - add this to the top of each checklist sheet?
    // var accessInfo = scopeData[i][2]; // future - add this to the top of each checklist sheet?
    var appendToName = " | checklist";
    
    var newSheetName = name + appendToName;
    
    var potentialSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(newSheetName);
    if (potentialSheet == null) {
      checklistTemplateSheet.copyTo(auditSpreadsheet).setName(newSheetName); // creates duplicate of checklist and rename
    }    
  }
  
  auditSpreadsheet.deleteSheet(checklistTemplateSheet);
}

function setAuditTypes() {  
  var checkListSheetHeaders = SpreadsheetApp.getActiveSheet().getRange("Checklist!1:1").getValues(); // first row of checklistSheet
  var auditTypesArray = []; // empty array for storing types of arrays
  
  for (var i=0; i < checkListSheetHeaders[0].length; i++) {
    var potentialAuditType = checkListSheetHeaders[0][i];
    
    if (potentialAuditType.includes("Audit - ")) {
      auditTypesArray.push(potentialAuditType.substring(8));
    }
  }
  
  var auditTypeCell = detailsSheet.getRange('B1'); // future - set to be relative to B column when A column that has "Audit Type"
  var auditTypeValidation = SpreadsheetApp.newDataValidation().requireValueInList(auditTypesArray).build();
  auditTypeCell.setDataValidation(auditTypeValidation);
}

function copyToAuditDirectory() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Copy Audit Spreadsheet', 'Name', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    var fileName = response.getResponseText();
    var auditSpreadsheetID = SpreadsheetApp.getActiveSpreadsheet().getId();
    //var folderID = ''; // set to ID of the folder ex. https://drive.google.com/drive/folders/<id of folder>
    var folder = DriveApp.getFolderById(folderID);

    DriveApp.getFileById(auditSpreadsheetID).makeCopy(fileName, folder);
  }
}

// from https://stackoverflow.com/questions/21229180/convert-column-index-into-corresponding-column-letter
function columnToLetter(column)
{
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}