# sheetsextra
Extra classes and methods to make scripting for Google Spreadsheets much easier.

To use, go to Resources > Library in the Google Sheets script editor and enter the following project key: MBMqhMPNTAcCr_ppXue1xbWEdFEUJsEaZ

```javascript
function getOrCreateSheet(spreadsheet, sheetName)
function renameSheet(spreadsheet, sheet, newName)
function renameSheetIfExists(spreadsheet, sheetName, newName)
function moveSheet(spreadsheet, sheet, newpos)
function moveSheetIfExists(spreadsheet, sheetName, newpos)
function duplicateSheet(spreadsheet, sheet)
function duplicateSheetIfExists(spreadsheet, sheetName)
function duplicateRenameSheet(spreadsheet, sheet, newName, renameOld)
function duplicateRenameSheetIfExists(spreadsheet, sheetName, newName, renameOld)
function deleteSheetIfExists(spreadsheet, sheetName)
function dataRangeWithoutHeader(sheet, rowHeaderSize)
function sortAll(sheet, sortOptions, rowHeaderSize)
function autoResizeAll(sheet)
function boldHeaders(sheet)
function freezeHeaders(sheet, rows, columns)

```

Sheetslibrary documentation here:

https://docs.google.com/macros/library/d/MBMqhMPNTAcCr_ppXue1xbWEdFEUJsEaZ/12
