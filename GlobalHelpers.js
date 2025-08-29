/**
 * Get sheet data as an array of objects.
 * Each object = {header1: value1, header2: value2, ...}
 *
 * @param {string} sheetName - Name of the sheet to read
 * @returns {Array<Object>} - Array of row objects
 */
function getDataAsObjects(sheet, startRow) {
  startRow = startRow || 1;
  const range = sheet.getRange(startRow, 1, sheet.getLastRow() - (startRow-1), sheet.getLastColumn());
  const values = range.getValues();
  
  const headers = values.shift(); // first row as headers
  const objects = values.map(row => {
    let obj = {};
    headers.forEach((header, i) => {
      obj[header] = row[i];
    });
    return obj;
  });
  
  return objects;
}

/**
 * Get sheet data as a dictionary (object) with keys from a specified column.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to read.
 * @param {string} keyField The name of the column to use as the dictionary key.
 * @param {number} headerRow The 1-based row index of the headers.
 * @returns {Object.<string, Object>} A dictionary where keys are from the specified column and values are the row objects.
 */
function getDataAsDict(sheet, keyField, headerRow) {
  // Get the headers from the specified header row.
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const keyFieldIndex = headers.indexOf(keyField);
  
  // Check if the key field exists in the headers.
  if (keyFieldIndex === -1) {
    console.error(`Key field "${keyField}" not found in headers.`);
    return {};
  }

  // Get the data, starting from the row after the headers.
  const numRows = sheet.getLastRow() - headerRow;
  if (numRows <= 0) {
    return {};
  }
  
  const dataRange = sheet.getRange(headerRow + 1, 1, numRows, sheet.getLastColumn());
  const dataValues = dataRange.getValues();
  
  // Build the dictionary.
  const dictionary = {};
  dataValues.forEach(row => {
    let obj = {};
    headers.forEach((header, i) => {
      obj[header] = row[i];
    });
    
    const key = row[keyFieldIndex];
    if (key !== undefined && key !== null && key !== '') {
      dictionary[key] = obj;
    }
  });
  
  return dictionary;
}

/**
 * Finds the first row that has a specified value in a given column.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to search.
 * @param {number} columnNumber The column number (1-based) to search in.
 * @param {string} valueToFind The value to look for.
 * @return {number} The 1-based row index, or -1 if not found.
 */
function findRowByValue(sheet, columnNumber, valueToFind) {
  // Get all values from the specified column.
  const columnValues = sheet.getRange(1, columnNumber, sheet.getLastRow()).getValues();
  
  // Iterate through the values to find the first match.
  for (let i = 0; i < columnValues.length; i++) {
    if (columnValues[i][0] === valueToFind) {
      // Return the 1-based row index.
      return i + 1;
    }
  }
  
  // Return -1 if the value is not found.
  return -1;
}

/**
 * Clears and updates a sheet with data from a list of objects.
 * * @param {Array<Object>} data - The array of objects to write to the sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to update.
 */
function updateSheetFromObjects(sheet, data) {
  if (!sheet) {
    console.error("Sheet reference is null or undefined.");
    return;
  }
  
  // Check if there's any data to write.
  if (!data || data.length === 0) {
    // If no data, just clear the sheet except for headers.
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
    }
    return;
  }
  
  const headers = Object.keys(data[0]);
  
  // Clear all content in the sheet, keeping the headers.
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
  }
  
  // Convert the array of objects into a 2D array of values.
  const values = data.map(obj => headers.map(header => obj[header]));
  
  // Write the new data to the sheet.
  sheet.getRange(2, 1, values.length, headers.length).setValues(values);
}