function importCSVFromGoogleDrive() {
  const fileName = "PolymerDispFrequency0.csv";
  const csvFile = DriveApp.getFilesByName(fileName).next();
  const csvData = Utilities.parseCsv(csvFile.getBlob().getDataAsString('ISO-8859-1'));

  const bufferSheetName = "BufferSheet0";
  const bufferSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(bufferSheetName) || SpreadsheetApp.getActiveSpreadsheet().insertSheet(bufferSheetName);
  bufferSheet.clear();

  const bufferRange = bufferSheet.getRange(1, 1, csvData.length, csvData[0].length);
  bufferRange.setValues(csvData);

  const targetSheetName = "PolymerDispFrequency0";
  const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(targetSheetName);
  if (!targetSheet) {
    throw new Error(`Sheet not found: ${targetSheetName}`);
  }

  const validityColumnIndex = getColumnIndexByHeader(targetSheet, "Validity");
  const targetData = getSheetData(targetSheet);
  const bufferData = getSheetData(bufferSheet);

  const newRows = getNewRows(bufferData, targetData, validityColumnIndex);

  if (newRows.length > 0) {
    appendRows(targetSheet, newRows);
  }

  deleteSheet(bufferSheet);
}

function getSheet(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) {
    throw new Error(`Sheet not found: ${name}`);
  }
  return sheet;
}

function getColumnIndexByHeader(sheet, header) {
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headerRow.findIndex(value => value === header);
}

function getSheetData(sheet) {
  const range = sheet.getDataRange();
  const values = range.getValues();
  const headerRowIndex = 0;
  return Array.from(values, (row, rowIndex) => {
    if (rowIndex === headerRowIndex) {
      return row;
    }
    return row.filter(value => value !== "");
  });
}

function getNewRows(bufferData, targetData, validityColumnIndex) {
  const uniqueIds = targetData.reduce((ids, row) => {
    if (row[validityColumnIndex] !== 100) {
      ids[row.join("|")] = true;
    }
    return ids;
  }, {});

  const newRows = bufferData.reduce((rows, row) => {
    if (row[validityColumnIndex] !== 100 && !uniqueIds[row.join("|")]) {
      rows.push(row);
      uniqueIds[row.join("|")] = true;
    }
    return rows;
  }, []);

  return newRows;
}

function appendRows(sheet, rows) {
  const startRow = sheet.getDataRange().getNumRows() + 1;
  const range = sheet.getRange(startRow, 1, rows.length, rows[0].length);
  range.setValues(rows);
}

function deleteSheet(sheet) {
  SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheet);
}
