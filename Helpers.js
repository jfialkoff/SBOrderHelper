/**
 * Get sheet data as an array of objects.
 * Each object = {header1: value1, header2: value2, ...}
 *
 * @param {string} sheetName - Name of the sheet to read
 * @returns {Array<Object>} - Array of row objects
 */
function getDataAsObjects(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const range = sheet.getDataRange();
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