  function importCSVFromGoogleDrive() {
  
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PolymerDispFrequency0");
     sheet.clearContents(); 
    var file2 = DriveApp.getFilesByName("PolymerDispFrequency0.csv").next();
    var csvData2 = Utilities.parseCsv(file2.getBlob().getDataAsString('ISO-8859-1'));
    var sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PolymerDispFrequency0");
      sheet2.getRange(1,1, csvData2.length, csvData2[0].length).setValues(csvData2);
  
  }
