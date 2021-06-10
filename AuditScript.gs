/*
  Audit Spreadsheet, Accessibility Audit Checklist
  Copyright (C) 2017-2021  Luis Garcia
  Contact: audit@garcialo.com
  URL: http://audit.garcialo.com
  GitHub Repo: https://github.com/garcialo/auditSpreadsheet
  
  This program is free software: you can redistribute it and/or modify
  it under the terms of the GNU General Public License as published by
  the Free Software Foundation, either version 3 of the License, or
  (at your option) any later version.
  
  This program is distributed in the hope that it will be useful,
  but WITHOUT ANY WARRANTY; without even the implied warranty of
  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
  GNU General Public License for more details.
  
  You should have received a copy of the GNU General Public License
  along with this program.  If not, see <https://www.gnu.org/licenses/>.
*/

var auditSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var checklistSheet = auditSpreadsheet.getSheetByName("Checklist");
var detailsSheet = auditSpreadsheet.getSheetByName("Details");
var scopeSheet = auditSpreadsheet.getSheetByName("Scope");

function onOpen() {
  setAuditTypes();
  createUI();
}

function createUI() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Audit')
    .addItem('Create checklists from Scope', 'createChecklists')
    .addItem('Create All Issues sheet', 'createAllIssuesSheet')
    .addItem('Add custom audit type', 'addCustomAuditType')
    .addToUi();
}

function createAllIssuesSheet() {
  // create list of all issues from all checklist sheets
  var allSheets = auditSpreadsheet.getSheets();
  var allIssuesArray = [];

  for (var i = 0; i < allSheets.length; i++) {
    var currentSheet = allSheets[i];
    if (currentSheet.getName().includes(" | checklist") && currentSheet.getRange('B1').getValue() === 'Status') { 
      var currentRange = currentSheet.getDataRange();
      var numRows = currentRange.getLastRow() - 1; // -1 because we're starting on the second row
      var numColumns = 5; // 
      
      var issuesRange = currentSheet.getRange(2,1, numRows, numColumns);
      var issuesValues = issuesRange.getValues();
      
      for (var j = 0; j < numRows; j++) {
        var statusCell = issuesValues[j][1];
        var issueCell = issuesValues[j][2];

        // determine when to add issues to the list
        /*
        - pass: do not add to list
        - n/a:  do not add to list
        - fail: add to list
        - incomplete and has notes:  add to list
        - notes/other and has notes: add to list
        */
        if ((statusCell === 'Fail') ||
            (statusCell === 'Check Incomplete' || statusCell === 'Notes / Other') && issueCell !== '') {
          var trackerCell = issuesValues[j][0];
          var descriptionCell = issuesValues[j][4];
          var currentSheetName = currentSheet.getName().slice(0,-12); // cut " | checklist" off of the sheet name

          var output = [trackerCell, statusCell, issueCell, descriptionCell , currentSheetName];
          allIssuesArray.push(output);
        }
      }
    }
  }
  
  // create All Issues sheet if the list of issues actually has issues
  if (allIssuesArray.length > 0) {
    var allIssuesSheet = '';
  
    var potentialSheet = auditSpreadsheet.getSheetByName('All Issues');
    if (potentialSheet == null) {
      allIssuesSheet = auditSpreadsheet.insertSheet('All Issues',1);
      allIssuesSheet.getRange(1, 1, 1, 5).setValues([["Issue Tracker", "Status", "Issue", "Description", "Page or Component"]]);
    }
    else {
      allIssuesSheet = auditSpreadsheet.getSheetByName("All Issues");
    }
  
    // freeze first row
    allIssuesSheet.setFrozenRows(1);
   
    // bold first row
    allIssuesSheet.getRange("'All Issues'!A1:E1").setFontWeight("bold");
  
    //set wrapping on Description
    var allIssuesDescriptionColumn = allIssuesSheet.getRange("D1:D"); // change this to Description column in All Issues
    allIssuesDescriptionColumn.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  
    // set font size on sheet
    var allIssuesEntireSheetRange = allIssuesSheet.getRange("'All Issues'!1:1000");
    allIssuesEntireSheetRange.setFontSize(12);
  
    // set font color for entire sheet
    allIssuesEntireSheetRange.setFontColor("#434343");

    // column sizes - setting non-auto column sizes
    allIssuesSheet.setColumnWidth(1, 130); // "Issue Tracker"
    allIssuesSheet.setColumnWidth(3, 130); // "Issues"
    allIssuesSheet.setColumnWidth(4, 510); // "Description"
    
    // add list of all issues to the All Issues sheet
    numRows = allIssuesArray.length;
    numColumns = 5;
    allIssuesSheet.getRange(2,1, numRows, numColumns).setValues(allIssuesArray);
    
    //set bg color, but only for data range
    var allIssuesDataRange = allIssuesSheet.getDataRange();
    allIssuesDataRange.setBackground("#FFF2CC");
  
    // column sizes - setting autoresized columns
    allIssuesSheet.autoResizeColumn(2); // "Status"
    allIssuesSheet.autoResizeColumn(5); // "Page or Component"
  
    // protect sheet
    allIssuesSheet.protect().setWarningOnly(true);
  }
  else {
    SpreadsheetApp.getUi().alert("No issues found in page/component checklists");
  }
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
    
    var potentialSheet = auditSpreadsheet.getSheetByName(newSheetName);
    if (potentialSheet == null) {
      checklistTemplateSheet.copyTo(auditSpreadsheet).setName(newSheetName); // creates duplicate of checklist and rename
    }    
  }
  
  // delete the template sheet since it's no longer needed
  auditSpreadsheet.deleteSheet(checklistTemplateSheet);
  
  // moving Checklist to the end since it's not needed as much now
  var lastSheetPosition = auditSpreadsheet.getNumSheets();
  checklistSheet.activate();
  auditSpreadsheet.moveActiveSheet(lastSheetPosition);
  
  // activating Scope - eventually should have links to the created sheets
  scopeSheet.activate();
}

function addCustomAuditType() {
  var ui = SpreadsheetApp.getUi();
  var newCustomAuditName = ui.prompt('New Custom Audit', 'Name', ui.ButtonSet.OK_CANCEL);

  checklistSheet.insertColumnAfter(14); // 14 = column N (Audit - Full)
  var newAuditColumnHeader = checklistSheet.getRange("O1")
  newAuditColumnHeader.setValue("Audit - " + newCustomAuditName.getResponseText());

  var fullAuditRange = checklistSheet.getRange(2, 14, checklistSheet.getMaxRows() - 1, 1);
  fullAuditRange.copyTo(checklistSheet.getRange(2, 15));
  checklistSheet.autoResizeColumn(15);

  setAuditTypes();
  newAuditColumnHeader.activate();
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

// look into finding a way to move the current file instead of just making a copy every single time
function copyToAuditFolder() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Copy Audit Spreadsheet', 'Name', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    var fileName = response.getResponseText();
    var auditSpreadsheetID = auditSpreadsheet.getId();
    var auditFolderID = ''; // set to ID of shared folder everyone should save Audits to ex. https://drive.google.com/drive/folders/<id of folder>
    var folder = DriveApp.getFolderById(auditFolderID);

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
