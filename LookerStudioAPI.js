function importCSVFromGoogleDrive() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("AirCompressors0");
    var file = DriveApp.getFilesByName("AirCompressors0.csv").next();
    var csvData = Utilities.parseCsv(file.getBlob().getDataAsString('ISO-8859-1'));
    var lastRow = sheet.getLastRow();
    var numRows = csvData.length;
    var numColumns = csvData[0].length;
    sheet.insertRowsAfter(lastRow, numRows);
    sheet.getRange(lastRow + 1, 1, numRows, numColumns).setValues(csvData);
  }
