function importCSVFromGoogleDrive() {
  var csvFile = DriveApp.getFilesByName("csvName.csv").next();
  var csvData = Utilities.parseCsv(csvFile.getBlob().getDataAsString('ISO-8859-1'));
  var bufferSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BufferSheet0");
  if (bufferSheet == null) {
    bufferSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("BufferSheet0");
  }
  bufferSheet.clearContents();
  bufferSheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);

  var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SheetName");
  var lastRow = targetSheet.getLastRow();
  var newData = getNewRows(bufferSheet, targetSheet, lastRow);

  addNewRows(targetSheet, lastRow, newData);
  SpreadsheetApp.getActiveSpreadsheet().deleteSheet(bufferSheet);
}

function getNewRows(bufferSheet, targetSheet, lastRow) {
  var bufferRange = bufferSheet.getDataRange();
  var bufferData = bufferRange.getValues();
  var targetRange = targetSheet.getRange(1, 1, lastRow, bufferData[0].length);
  var targetData = targetRange.getValues();
  var existingData = {};

  for (var i = 0; i < targetData.length; i++) {
    var key = targetData[i].toString();
    existingData[key] = true;
  }

  var newData = [];
  for (var i = 0; i < bufferData.length; i++) {
    var key = bufferData[i].toString();
    if (!existingData[key]) {
      newData.push(bufferData[i]);
      existingData[key] = true;
    }
  }

  return newData;
}

function addNewRows(targetSheet, lastRow, newData) {
  var numRows = newData.length;
  var numColumns = newData[0].length;
  if (numRows > 0) {
    targetSheet.insertRowsAfter(lastRow, numRows);
    targetSheet.getRange(lastRow + 1, 1, numRows, numColumns).setValues(newData);
  }
}
