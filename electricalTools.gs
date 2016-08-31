function onOpen() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var menu1Entries = [{name: "Create New Project", functionName: "createBomSheet"},
                      {name: "Reset Project", functionName: "reset"},
                      // {name: "Export Sheets as .xlsx", functionName: "exportAsxlsx"},
                      {name: "Backup Project", functionName: "saveBackup"},
                      {name: "Recover Backup", functionName: "recoverBackup"},
                       ];
  var menu2Entries = [{name: "Sequence - Process BOM (run all)", functionName: "processBom"},
                      {name: "Filter BOM", functionName: "filterPullSection"},
                      {name: "Unhide Filtered Parts", functionName: "unhideAllRows"},
                      {name: "Apply Color Gradient", functionName: "colorSections"},
                      {name: "List Filtered Parts", functionName: "seeFiltered"}
                      ];
  var menu3Entries = [
                      {name: "Create Epicor BOM", functionName: "createEpicorSheet"},
                      {name: "Create Part Serial Form", functionName: "createPartSNSheet"}
                      ];
  ss.addMenu("Project", menu1Entries);
  ss.addMenu("E-BOM Functions", menu2Entries);
  ss.addMenu("Create Electrical Worksheets", menu3Entries);
}

function createSheet(sheetName, copyMachInfoBool){
  // Make a new sheet with a sheet name argument. Optional true argument will copy the header info from the main BOM page.
  var ss = SpreadsheetApp.getActive();
  try {
    var sheet = ss.insertSheet(sheetName);
  } 
  catch (e) {
    var sheet = ss.getSheetByName(sheetName);
  }
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 75);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 75);
  if(copyMachInfoBool !== undefined && copyMachInfoBool === true){
    copyMachineInfo(sheetName);
  }
}

function copyMachineInfo(sheetName){
  // Copy header info from EBOM sheet
  var headerRows = 3;
  var headerCols = 5;
  var ss = SpreadsheetApp.getActive();
  var srcSheet = ss.getSheetByName("EBOM");
  var dstSheet = ss.getSheetByName(sheetName);
  var srcRange = srcSheet.getRange(1, 1, headerRows, headerCols);
  var srcData = srcRange.getValues();
  var destRange = dstSheet.getRange(1, 1, headerRows, headerCols);
  var letterArray = ["A", "B", "C", "D", "E"];
  srcRange.copyFormatToRange(dstSheet, 1, headerCols, 1, headerRows); 
  destRange.setValues(srcData);
  // Set formulas in header cells to copy header info from main EBOM page.
  for(var i = 1; i <= letterArray.length; i++){
    destRange.getCell(3,i).setFormula("=EBOM!" + letterArray[i - 1] + "3");
  }
}

function getSerialNumber(){
  // Regex to make serial number out of electrical print number
  var ss = SpreadsheetApp.getActive();
  var year = new Date().getUTCFullYear().toString().slice(2);
  var printNumber = ss.getRangeByName("printNumber").getValue();
  Logger.log(printNumber);
  var patt = /-P.*$/
  var machNumberNoPrg = printNumber.replace(patt, "");
  var serialNumber = year + "-" + machNumberNoPrg;
  return serialNumber;
}

function filterPullSection(returnSkipped){
  // Manual spreadsheet filter, return the hidden rows so user can make sure the script is working correctly (not implemented yet)
  var ss = SpreadsheetApp.getActive();
  var bomSheet = ss.getSheetByName("EBOM");
  var bomLastRow = bomSheet.getLastRow();
  var firstItemRow = 5;
  var pullSectionCol = 3;
  var bomRange = bomSheet.getRange(firstItemRow, 1, bomLastRow - firstItemRow + 1, 4)
  var bomData = bomRange.getValues();
  var filteredArray = [];
  var skippedArray = [];
  for (var i = 0; i < bomData.length; i++){
    var rowArr = bomData[i];
    var pullSection = rowArr[pullSectionCol];
    if (pullSection === "CONTROL" || pullSection === "MACHINE" || pullSection === "PANEL"){
      filteredArray.push(rowArr);
    }
    else{
      bomSheet.hideRows(i + firstItemRow);
      // Ignore rows without part numbers or the pasted Type row
      if(!(rowArr[0] === undefined || rowArr[0] === "" || rowArr[0] === "Type")){
        // Push the part number that is skipped
        skippedArray.push(rowArr[0]);
      }
    }
  }
  // Returns skipped array if argument is true
  if(returnSkipped === true){
    return skippedArray;
  }
  return filteredArray;
}

function seeFiltered(){
  var skippedParts = filterPullSection(true);
  SpreadsheetApp.getUi().alert("Filtered Parts: \n" + skippedParts.join("\n"));
}

function indexSections(){
  // Get indices of each pull section for further operations (can add color to each section, generate separate BOMs for each section, etc)
  var ss = SpreadsheetApp.getActive();
  var bomSheet = ss.getSheetByName("EBOM");
  var bomLastRow = bomSheet.getLastRow();
  var bomRange = bomSheet.getRange(4, 1, bomLastRow - 3, 5);
  // Create index array with first index of zero
  var indexArray = [0];
  var checkValue = "";
  for(var i = 1; i < bomLastRow; i++){
    var bomRow = bomSheet.getRange(3 + i, 1, 1, 5);
    var newValue = bomRow.getCell(1,4).getValue();
    if(newValue !== checkValue && (newValue === "CONTROL" || newValue === "MACHINE" || newValue === "PANEL")){
      checkValue = newValue;
      indexArray.push(i - 1);
    }
  }
  colorSections(indexArray);
}

function colorSections(indexArr){
  // Get each section's range and apply an alternating gray background color for easy visual differentiation
  var sheet = SpreadsheetApp.getActive().getSheetByName("EBOM");
  var lastRow = sheet.getLastRow();
  var grayBool = true;
  var firstItem = 4;
  for(var i = 0; i < indexArr.length - 1; i++){
    var range = sheet.getRange(firstItem + indexArr[i], 1, indexArr[i + 1] - indexArr[i], 5);
    if(grayBool){
      range.setBackground("#d9d9d9");
    }
    else{
      range.setBackground("#ffffff");
    }
    grayBool = !grayBool;
  }
  sheet.getRange(1, 1, lastRow, 5).setBorder(true, true, true, true, true, true);
}

function createBomSheet(){
  // Create BOM template page from reset spreadsheet
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActive();
  var printNumRange = ss.getRangeByName("printNumber");
  var jobNumRange = ss.getRangeByName("jobNumber");
  if(printNumRange.isBlank() || jobNumRange.isBlank()){
    return ui.alert("You have not entered a valid print number and job number (make sure you are not editing an input cell). Please try again.");
  }
  var sheetName = "EBOM";
  createSheet(sheetName);
  var sheet = ss.getSheetByName(sheetName);
  //
  ss.setActiveSheet(sheet);
  var lastRow = sheet.getLastRow();
  var sn = getSerialNumber();
  var date = new Date();
  var formattedDate = (date.getMonth() + 1) + "/" + date.getDate() + "/" + date.getFullYear();
  sheet.getRange(1, 1, 1, 5).merge();
  sheet.getRange(1, 1, 1, 5).setFontSize(14);
  var globalHeaderRange = sheet.getRange(1, 1, 3, 5);
  // Set cell format of job number to string to account for zeros prefix 000###
  var jobNumberRange = sheet.getRange(3, 3).setNumberFormat("@");
  // Sheet formatting and static labels
  globalHeaderRange.setValues([["ELECTRICAL BOM FOR SERIAL NUMBER: " + sn, "", "", "", ""],
                                       ["MACHINE SERIAL NUMBER", "PRINT NUMBER", "JOB #", "DATE", "INIT"],
                                       [sn, printNumRange.getValue(), jobNumRange.getValue(), formattedDate, ""]]);
  var bomHeaderRange = sheet.getRange(4, 1, 2, 5);
  var headerArr = [["KVAL PART NUMBER", "DESCRIPTION", "QTY", "PULL", "INIT"], 
                   ["SELECT THIS CELL & PASTE", "Type description", "Amount", "User Setting 11", "Init"]];
  bomHeaderRange.setValues(headerArr);
  // Format gloabl header ranges
  sheet.getRange(1, 1, 5, 5).setHorizontalAlignment("center");
  sheet.getRange(1, 1, 1, 5).setBackground("#6aa74f");
  sheet.getRange(2, 1, 1, 5).setBackground("#b5d5a7");
  sheet.getRange(2, 1, 1, 5).setFontWeight("bold");
  sheet.getRange(3, 1, 1, 5).setBackground("#d9ead3");
  sheet.getRange(4, 1, 1, 5).setBackground("#cccccc");
  sheet.getRange(4, 1, 1, 5).setFontWeight("bold");
  sheet.getRange(5, 1, 1, 5).setBackground("#efefef");
  var pasteHereCellRange = sheet.getRange(5, 1);
  pasteHereCellRange.setBackground("yellow");
  // Add blank rows for writing in missing parts
  var maxRowsPrinted = 91;
  sheet.getRange(1, 1, Math.max(maxRowsPrinted, lastRow), 1).setNumberFormat('@');
  sheet.getRange(1, 3, Math.max(maxRowsPrinted, lastRow), 1).setHorizontalAlignment("center");
  sheet.getRange(1, 1, Math.max(maxRowsPrinted, lastRow), 5).setBorder(true, true, true, true, true, true);
}

function createEpicorArray(bomArr){
  // Create array with empty elements for column formatting for paste insert into database
  var blankColumns = [0, 2, 4, 5];
  var count = 0;
  bomArr.map(function(bomElement){
    for(var i in blankColumns){
      bomElement.splice(blankColumns[i] + count, 0, "");
    }
    bomElement.splice(-1);
  });
  return bomArr;
}      
         
function createEpicorSheet(){
  // Use epicor array to build epicor sheet
  var firstItem = 4;
  var sheetName = "ImportToEpicor";
  createSheet(sheetName, true);
  var filteredArr = filterPullSection();
  var epicorArr = createEpicorArray(filteredArr);
  var numRows = epicorArr.length;
  var numCols = epicorArr[0].length;
  var ss = SpreadsheetApp.getActive();
  var range = ss.getSheetByName(sheetName).getRange(firstItem, 1, numRows, numCols);
  var partNumberRange = ss.getSheetByName(sheetName).getRange(firstItem, 2, numRows, 1);
  // Sheet formatting
  partNumberRange.setNumberFormat('@');
  range.setValues(epicorArr);
  range.setBackground('#efefef');
}

function createPartSNArray(bomArr){
  // Create array for part numbers that require documentation of manufacturer's serial number
  var resultArray = [];
  var matchArray = ["BEC-AX5", "BEC-CX2020", "BEC-CX90", "FRE7"];
  bomArr.map(function(bomElement){
    for(var i = 0; i < matchArray.length; i++){
      if(bomElement[0].search(matchArray[i]) !== -1){
        var count = bomElement[2];
        for(var j = 1; j <= count; j++){
          resultArray.push([bomElement[0], "", j, count, ""]);
        }
      }
    }
  });
  return resultArray
}

function createPartSNSheet(){
  // Use part SN array to create sheet, multiple quantities of matching parts will write the instance number and total number to the sheet.
  var ui = SpreadsheetApp.getUi();
  var sheetName = "PartSerialNumbers";
  createSheet(sheetName, true);
  var filteredArr = filterPullSection();
  var partSNArr = createPartSNArray(filteredArr);
  var rowStart = 4;
  var numRows = partSNArr.length;
  var numCols = partSNArr[0].length;
  // Early return if no matching parts
  if(numRows === 0){
    return false;
  }
  // Sheet formatting
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(sheetName);
  var range = sheet.getRange(rowStart + 1, 1, numRows, numCols);
  range.setValues(partSNArr);
  var headerRange = sheet.getRange(rowStart, 1, 1, numCols);
  var headerArray = [["KVAL PART NUMBER", "SERIAL NUMBER", "INSTANCE", "QUANTITY", "INIT"]];
  var borderRange = sheet.getRange(1, 1, numRows, numCols);
  headerRange.setBackground("efefef");
  headerRange.setValues(headerArray);
  headerRange.setHorizontalAlignment("center");
  headerRange.setBorder(true, true, true, true, true, true);
  range.setBorder(true, true, true, true, true, true);
  borderRange.setBorder(true, true, true, true, true, true);
}


function processBom(){
  // Run BOM Processing Sequence
  filterPullSection();
  indexSections();
  saveBackup();
  createPartSNSheet();
  createEpicorSheet();
}

function reset(bypass){
  // Reset the sheet
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets();
  var ui = SpreadsheetApp.getUi();
  ss.getSheetByName("BACKUP").hideSheet();
  // Bypass ui alert if true argument
  if(bypass !== undefined && bypass === true){
    // Delete all sheets besides UI and backup
    for(var i = 0; i < sheets.length; i++){
      if(sheets[i].getSheetName() !== "User Interface" && sheets[i].getSheetName() !== "BACKUP"){
        ss.deleteSheet(sheets[i]);
      }
    }
    // Bypass alert
    return true;
  }
  var response = ui.alert("Are you sure you want to reset?", ui.ButtonSet.YES_NO);
  if(response == ui.Button.YES){
    for(var i = 0; i < sheets.length; i++){
      if(sheets[i].getSheetName() !== "User Interface" && sheets[i].getSheetName() !== "BACKUP"){
        ss.deleteSheet(sheets[i]);
      }
    }
  }
}

function clearInputFields(){
  // Clear input fields
  var ss = SpreadsheetApp.getActive();
  ss.getRangeByName("printNumber").clearContent();
  ss.getRangeByName("jobNumber").clearContent();
}

//function exportAsxlsx() {
//  var spreadsheet   = SpreadsheetApp.getActiveSpreadsheet();
//  var spreadsheetId = spreadsheet.getId()
//  var file          = Drive.Files.get(spreadsheetId);
//  var url           = file.exportLinks[MimeType.MICROSOFT_EXCEL];
//  var token         = ScriptApp.getOAuthToken();
//  var response      = UrlFetchApp.fetch(url, {
//    headers: {
//      'Authorization': 'Bearer ' +  token
//    }
//  });
//  var fileName = Browser.inputBox("Save xlsx file as:")
//  var blobs   = response.getBlob();
//  var folder = DriveApp.getFoldersByName('Exports');
//  if(folder.hasNext()) {
//    var existingPlan1 = DriveApp.getFilesByName(fileName + '.xlsx');
//    if(existingPlan1.hasNext()){
//      var existingPlan2 = existingPlan1.next();
//      var existingPlanID = existingPlan2.getId();
//      Drive.Files.remove(existingPlanID);
//    }
//  } else {
//    folder = DriveApp.createFolder('Exports');
//  }
//  folder = DriveApp.getFoldersByName('Exports').next();
//  folder.createFile(blobs).setName(fileName + '.xlsx');
//}

function unhideAllRows(){
  var sheet = SpreadsheetApp.getActive().getSheetByName("EBOM");
  var fullRange = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
  sheet.unhideRow(fullRange);
}
                 

function saveBackup(){
  // Save last BOM page data to hidden sheet
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("EBOM");
  var ui = SpreadsheetApp.getUi();
  var maxRows = sheet.getMaxRows();
  var maxCols = sheet.getMaxColumns();
  var bomData = sheet.getRange(1, 1, maxRows, maxCols).getValues();
  var sheetName = "BACKUP";
  try {
    var backupSheet = ss.insertSheet(sheetName);
    backupSheet.hideSheet();
    backupSheet.getRange(1, 1, maxRows, maxCols).setNumberFormat("@")
    backupSheet.getRange(1, 1, maxRows, maxCols).setValues(bomData);
  } 
  catch (e) {
    var backupSheet = ss.getSheetByName(sheetName);
    var response = ui.alert("Discard your previous backup?", ui.ButtonSet.YES_NO);
    if(response == ui.Button.YES){
      backupSheet.hideSheet();
      backupSheet.getRange(1, 1, maxRows, maxCols).setNumberFormat("@")
      backupSheet.getRange(1, 1, maxRows, maxCols).setValues(bomData);
    }
  } 
}

function recoverBackup(){
  // Write BOM array from backup sheet to current BOM sheet
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var backupSheet = ss.getSheetByName("BACKUP");
  var response = ui.alert("Overwrite your current data for recovery?", ui.ButtonSet.YES_NO);
  if(response == ui.Button.YES){
    reset(true);
    var backupData = backupSheet.getRange(1, 1, backupSheet.getMaxRows(), backupSheet.getMaxColumns()).getValues();
    createBomSheet();
    ss.getSheetByName("EBOM").getRange(1, 1, backupSheet.getMaxRows(), backupSheet.getMaxColumns()).setValues(backupData);
  }
}
