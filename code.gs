/**
* Attempts to return a sheet belonging to a spreadsheet by name. If it doesn't exist the sheet will be created and returned.
*
* @param {Spreadsheet} spreadsheet the spreadsheet object
* @param {String} sheetName the name of the sheet to get or create
* @return {Sheet} the sheet object
*/
function getOrCreateSheet(spreadsheet, sheetName){
  if (spreadsheet.getSheetByName(sheetName) == null) {
    spreadsheet.insertSheet(sheetName);
  }     
  return spreadsheet.getSheetByName(sheetName);
}

/**
* Rename any sheet, not just the active one as Google Sheets normally only allows
*
* @param {Spreadsheet} spreadsheet the spreadsheet object
* @param {Sheet} sheet the sheet object to rename
* @param {String} sheet the new name for the sheet
*/
function renameSheet(spreadsheet, sheet, newName){
  var activeSheet = spreadsheet.getActiveSheet();
  spreadsheet.setActiveSheet(sheet);
  spreadsheet.renameActiveSheet(newName);
  if (activeSheet){
    spreadsheet.setActiveSheet(activeSheet);
  }
}


/**
* Attempts to rename a sheet by it's current name if it exists.
*
* @param {Spreadsheet} spreadsheet the spreadsheet object
* @param {String} sheetName the name of the sheet to rename
* @param {String} newName the new name for the sheet
* @return {Boolean} true if the sheet was found and renamed, false otherwise
*/
function renameSheetIfExists(spreadsheet, sheetName, newName){
  var sheet = spreadsheet.getSheetByName(sheetName); 
  if (sheet){
    renameSheet(spreadsheet, sheet, newName);
    return true;
  }
  else{
    return false;
  }
}

/**
* Move any sheet to a new position, not just the active one as Google Sheets normally only allows
*
* @param {Spreadsheet} spreadsheet the spreadsheet object
* @param {Sheet} sheet the sheet object to move
* @param {Integer} newpos the position to move the sheet to
*/
function moveSheet(spreadsheet, sheet, newpos){
  var activeSheet = spreadsheet.getActiveSheet();
  spreadsheet.setActiveSheet(sheet);
  spreadsheet.moveActiveSheet(newpos);
  if (activeSheet){
    spreadsheet.setActiveSheet(activeSheet);
  }
}

/**
* Attemts to move a sheet by name if it exists.
*
* @param {Spreadsheet} spreadsheet the spreadsheet object
* @param {string} sheetName the name of the sheet to move
* @param {Integer} newpos the position to move the sheet to
* @return {boolean} true if the sheet was found and moved, false otherwise
*/
function moveSheetIfExists(spreadsheet, sheetName, newpos){
  var sheet = spreadsheet.getSheetByName(sheetName); 
  if (sheet){
    moveSheet(spreadsheet, sheet, newpos)
    return true;
  }
  else{
    return false;
  }
}

/**
* Duplicate any sheet, not just the active one as Google Sheets normally only allows
*
* @param {Spreadsheet} spreadsheet the spreadsheet object
* @param {Sheet} sheet the sheet object to duplicate
* @return {Sheet} the new sheet
*/
function duplicateSheet(spreadsheet, sheet){
  var activeSheet = spreadsheet.getActiveSheet();
  spreadsheet.setActiveSheet(sheet);
  var newSheet = spreadsheet.duplicateActiveSheet();
  if (activeSheet){
    spreadsheet.setActiveSheet(activeSheet);
  }
  return newSheet;
}


/**
* Duplicate a sheet by name if it exists.
*
* @param {Spreadsheet} spreadsheet the spreadsheet object
* @param {string} sheetName the name of the sheet to duplicate
* @return {boolean} true if the sheet was found and moved, false otherwise
*/
function duplicateSheetIfExists(spreadsheet, sheetName){
  var sheet = spreadsheet.getSheetByName(sheetName); 
  if (sheet){
    duplicateSheet(spreadsheet, sheet);
    return true;
  }
  else{
    return false;
  }
}

/**
* Duplicate a sheet and rename either the old or new sheet
*
* @param {Spreadsheet} spreadsheet the spreadsheet object
* @param {Sheet} sheet the sheet object to duplicate
* @param {String} newName the new name of the sheet that will be renamed
* @param {Boolean} renameOld if true, renames the originally duplicated sheet, if false renames the newly created sheet
* @return {Sheet} the new sheet
*/
function duplicateRenameSheet(spreadsheet, sheet, newName, renameOld){
  var newSheet = duplicateSheet(spreadsheet, sheet);
  var renameNew = renameOld || false; 
  if (renameOld){
    renameSheet(spreadsheet, sheet, newName);
  }
  else{
    renameSheet(spreadsheet, newSheet, newName);
  }
  return newSheet;
}


/**
* Duplicate and renames a sheet by name if it exists.
*
* @param {Spreadsheet} spreadsheet the spreadsheet object
* @param {String} sheetName the name of the sheet to duplicate
* @param {String} newName the new name of the sheet that will be renamed
* @param {Boolean} renameOld if true, the originally duplicated sheet; if false renames the newly created sheet. Default is false.
* @return {Sheet} the new sheet or false if the sheet was not found
*/
function duplicateRenameSheetIfExists(spreadsheet, sheetName, newName, renameOld){
  var sheet = spreadsheet.getSheetByName(sheetName); 
  if (sheet){
    var renameNew = renameOld || false;
    var newSheet = duplicateRenameSheet(spreadsheet, sheetName, newName, renameOld);
    return newSheet;
  }
  else{
    return false;
  }
}

/**
* Attemts to delete a sheet by name if it exists.
*
* @param {Spreadsheet} spreadsheet the spreadsheet object
* @param {string} sheetName the name of the sheet to get or create
* @return {boolean} true if the sheet was found and deleted, false otherwise
*/
function deleteSheetIfExists(spreadsheet, sheetName){
  var sheet = spreadsheet.getSheetByName(sheetName); 
  if (sheet){
    spreadsheet.deleteSheet(sheet);
    return true;
  }
  else{
    return false;
  }
}

/**
* Returns a data range without the specified number of header rows
*
* @param {Sheet} sheet the sheet object
* @param {Integer} rowHeader the number of header rows. If not set a value of 1 is used.
* @return {Range} the data range
*/
function dataRangeWithoutHeader(sheet, rowHeaderSize){
  var rowHeaderSize = ((rowHeaderSize -1) || 0) + 1;
  if (sheet.getLastRow() <= rowHeaderSize){
      return false
  }  
  else{
      var rows = sheet.getRange(rowHeaderSize + 1, 1, sheet.getLastRow() - 1, sheet.getLastColumn());  
      return rows
  }
}


/**
* Sorts the full data range without the specified number of header rows
*
* @param {Sheet} sheet the sheet object
* @param {sortSpecObj} sortOptions the sort options as accepted by the Google Sheets sort.range method
* @param {Integer} rowHeaderSize the number of header rows. If not set a value of 0 is used.
* @return {Range} the newly sorted data range
*/
function sortAll(sheet, sortOptions, rowHeaderSize){
  var rowHeaderSize = rowHeaderSize || 0;
  var dataRange = dataRangeWithoutHeader(sheet, rowHeaderSize);
  if (dataRange){
    return dataRange.sort(sortOptions);
  }
  else{
    return false;
  }
}

/**
* Auto-resizes all the columns in a sheet
*
* @param {Sheet} sheet the sheet object
*/
function autoResizeAll(sheet){ 
  var lastColumn = sheet.getLastColumn();
  for ( var i = 1; i <= lastColumn; i++ ) {
    sheet.autoResizeColumn(i);
  }
}

/**
* Bolds all cells in Column A and Row 1
*
* @param {Sheet} sheet the sheet object
*/
function boldHeaders(sheet){
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight("bold");
  sheet.getRange(1, 1, sheet.getLastRow(), 1).setFontWeight("bold");
}

/**
* Freeze headers (1 row and 1 column by default)
*
* @param {Sheet} sheet the sheet object
* @param {Integer} rows the number of rows to freeze. Optional. Set to 1 if not given.
* @param {Integer} columns the number of columns to freeze. Optional. Set to 1 if not given
*/ 
function freezeHeaders(sheet, rows, columns){
  var rows = rows || 1;
  var columns = columns || 1;
  sheet.setFrozenRows(rows);
  sheet.setFrozenColumns(columns);
}

/**
* Converts a 2d list of values to a single array
*
* @param {Array} values2d 2d list of values
* @return {Array} values1d 1d list of values
*/ 
function to1d(values2d){
    var values1d = [];
    for(var i = 0; i < values2d.length; i++){
      values1d = values1d.concat(values2d[i]);
    }
    return values1d;
}

/**
* Sends script logfile to an e-mail address
*
* @param {String} mailto E-mail address to send logfile to
* @param {String} description Description included in e-mail subject
*/ 
function sendLogFile(mailto, description){
  date = new Date(),
    formattedDate = [date.getMonth()+1,
                     date.getDate(),
                     date.getFullYear()].join('/')+' '+
                       [date.getHours(),
                        date.getMinutes(),
                        date.getSeconds()].join(':');
  var subject = 'Logfile for '+ description + ' run at ' + formattedDate;
  var mailContents = Logger.getLog()
  try{
    MailApp.sendEmail(mailto, subject, mailContents);
  }
  catch(e){
    Logger.log(e); 
  }
}
