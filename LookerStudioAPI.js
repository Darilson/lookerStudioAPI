function importCSVFromGoogleDrive() {
    const fileName = "PolymerDispFrequency0.csv";
    Logger.log(`Importing file: ${fileName}`);
    const csvFile = DriveApp.getFilesByName(fileName).next();
    const csvData = Utilities.parseCsv(csvFile.getBlob().getDataAsString('ISO-8859-1'));

    const bufferSheetName = "BufferSheet0";
    const bufferSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(bufferSheetName) || SpreadsheetApp.getActiveSpreadsheet().insertSheet(bufferSheetName);
    bufferSheet.clear();
    Logger.log(`Created buffer sheet: ${bufferSheetName}`);

    const bufferRange = bufferSheet.getRange(1, 1, csvData.length, csvData[0].length);
    bufferRange.setValues(csvData);
    Logger.log(`Copied data to buffer sheet: ${bufferSheetName}`);

    const targetSheetName = "PolymerDispFrequency0";
    const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(targetSheetName);
    if (!targetSheet) {
        throw new Error(`Sheet not found: ${targetSheetName}`);
    }
    Logger.log(`Target sheet found: ${targetSheetName}`);

    const validityColumnIndex = getColumnIndexByHeader(targetSheet, "Validity");

    const targetData = getSheetData(targetSheet);
    Logger.log(`Target sheet data obtained`);

    const bufferData = getSheetData(bufferSheet);
    Logger.log(`Buffer sheet data obtained`);

    const newRows = getNewRows(bufferData, targetData, validityColumnIndex);

    if (newRows.length > 0) {
        appendRows(targetSheet, newRows);
    }

    deleteSheet(bufferSheet);
    Logger.log(`Buffer sheet deleted: ${bufferSheetName}`);
}

function getSheet(name) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
    if (!sheet) {
        throw new Error(`Sheet not found: ${name}`);
    }
    Logger.log(`Sheet found: ${name}`);
    return sheet;
}

function getColumnIndexByHeader(sheet, header) {
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const columnIndex = headerRow.findIndex(value => value === header);
    Logger.log(`Column index for header "${header}": ${columnIndex}`);
    return columnIndex;
}

function getSheetData(sheet) {
    const range = sheet.getDataRange();
    const values = range.getValues();
    const headerRowIndex = 0;
    const sheetData = Array.from(values, (row, rowIndex) => {
        if (rowIndex === headerRowIndex) {
            return row;
        }
        return row.filter(value => value !== "");
    });
    Logger.log(`Sheet data obtained for sheet: ${sheet.getName()}`);
    return sheetData;
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
    Logger.log(`New rows obtained: ${newRows.length}`);
    return newRows;
}

function appendRows(sheet, rows) {
    const startRow = sheet.getDataRange().getNumRows() + 1;
    const range = sheet.getRange(startRow, 1, rows.length, rows[0].length);
    range.setValues(rows);
    Logger.log(`New rows appended to target sheet: ${sheet.getName()}`);
}

function deleteSheet(sheet) {
    SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheet);
}
