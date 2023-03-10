function importCSVFromGoogleDrive() {
  var csvFile = DriveApp.getFilesByName("PolymerDispFrequency0.csv").next();
  var csvData = Utilities.parseCsv(csvFile.getBlob().getDataAsString('ISO-8859-1'));
  var bufferSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BufferSheet0");
  if (!bufferSheet) {
    bufferSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("BufferSheet0");
  }
  bufferSheet.clearContents();
  bufferSheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);

  var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PolymerDispFrequency0");
  if (!targetSheet) {
    throw new Error("Sheet not found: PolymerDispFrequency0");
  }

  var targetLastRow = targetSheet.getLastRow();
  var bufferLastRow = bufferSheet.getLastRow();
  if (bufferLastRow > targetLastRow) {
    var newRows = bufferSheet.getRange(targetLastRow + 1, 1, bufferLastRow - targetLastRow, csvData[0].length).getValues();
    targetSheet.getRange(targetLastRow + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
  }

  SpreadsheetApp.getActiveSpreadsheet().deleteSheet(bufferSheet);
}
