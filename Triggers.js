/**
 * A simple onEdit trigger that checks if the edited cell is in the
 * correct column and then calls the main function to update the dropdown.
 *
 * @param {Object} e The event object containing information about the edit.
 */
function onEdit(e) {
  console.log("Starting onEdit");
  // Exit if the event object is undefined, which happens when running manually.
  if (!e) {
    return;
  }
  
  // Get the active sheet and the range that was edited.
  const sheet = e.source.getActiveSheet();
  const editedRange = e.range;

  // Exit if the edited sheet is not the 'Order' sheet.
  if (sheet.getName() !== ORDER_SHEET_NAME) {
    return;
  }

  // Get the headers from the first row to find column indices.
  const headerRow = findRowByValue(sheet, 1, "Item");
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const itemColumn = headers.indexOf(ITEM_HEADER) + 1;

  // Exit if the edited column is not the "Item" column.
  if (editedRange.getColumn() !== itemColumn) {
    return;
  }
  
  // Call the main function to handle the dropdown update for the edited row.
  updatePathDropdown(sheet, editedRange.getRow());
}

/**
 * Updates the data validation for the "Path" column in a specific row
 * based on the "Item" selected in the same row.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The active sheet object.
 * @param {number} editedRow The row number that was edited.
 */
function updatePathDropdown(sheet, editedRow) {
  // Get the headers from the current sheet to find column indices.
  console.log("Starting updatePathDropdown");
  const headerRow = findRowByValue(sheet, 1, "Item");
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  console.log(headers);
  const pathColumn = headers.indexOf(PATH_HEADER) + 1;
  const inventoryIdColumn = headers.indexOf(INVENTORY_ID_HEADER) + 1;

  // Get the Inventory ID for the edited row.
  const inventoryId = sheet.getRange(editedRow, inventoryIdColumn).getValue();
  const pathCell = sheet.getRange(editedRow, pathColumn);
  pathCell.clearDataValidations();
  pathCell.setValue("Loading...");

  // If the Inventory ID is empty, clear the dropdown and exit.
  if (!inventoryId) {
    pathCell.setValue('Choose an item');
    return;
  }

  // Get the 'Paths' sheet and its data.
  const pathsSheet = sheet.getParent().getSheetByName(PATHS_SHEET_NAME);
  if (!pathsSheet) {
    console.error("The 'Paths' sheet was not found. Please ensure the sheet name is correct.");
    return;
  }

  // Find the headers in the Paths sheet.
  const pathsHeaders = pathsSheet.getRange(1, 1, 1, pathsSheet.getLastColumn()).getValues()[0];
  const pathsInventoryIdColumn = pathsHeaders.indexOf(INVENTORY_ID_HEADER) + 1;
  const ohLabelColumn = pathsHeaders.indexOf(OH_LABEL_HEADER) + 1;
  const pathsData = pathsSheet.getRange(2, 1, pathsSheet.getLastRow() - 1, pathsSheet.getLastColumn()).getValues();

  // Filter the Paths data to find matching 'OH Label' values.
  const pathOptions = pathsData
    .filter(row => row[pathsInventoryIdColumn - 1] === inventoryId)
    .map(row => row[ohLabelColumn - 1]);

  // Set the new data validation rule for the "Path" cell.
  if (pathOptions.length > 0) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(pathOptions)
      .setAllowInvalid(false) // Prevents users from entering values not in the list.
      .build();
    pathCell.setDataValidation(rule);

    // Automatically select the first option that starts with an asterisk.
    const defaultOption = pathOptions.find(option => String(option).startsWith('*'));
    if (defaultOption) {
      pathCell.setValue(defaultOption);
    }
  } else {
    // If no matching paths are found, clear any existing validation.
    pathCell.clearDataValidations();
  }
}

/**
 * A test function to manually run the updatePathDropdown function.
 * This is useful for debugging without having to manually edit a cell.
 */
function testUpdatePathDropdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName(ORDER_SHEET_NAME);
  
  // Choose a row number to test.
  const testRow = findRowByValue(orderSheet, 1, "Item") + 1;; 

  if (orderSheet) {
    console.log(`Testing dropdown update for row ${testRow}...`);
    updatePathDropdown(orderSheet, testRow);
    console.log("Test complete. Check the 'Path' cell in the specified row.");
  } else {
    console.error("Order sheet not found. Please make sure the sheet name is correct.");
  }
}
