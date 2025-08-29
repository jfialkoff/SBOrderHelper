function validateOrderRows() {
    const orderSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDER_SHEET_NAME);
    const headerRowIdx = findRowByValue(orderSheet, 1, "Item");
    const orderItemVals = getDataAsObjects(orderSheet, headerRowIdx);
    const pathVals = getPathData();
  
    const issues = [];
    orderItemVals.forEach(orderRow => {
      if(orderRow['Inventory ID'] === '')
        return;
      const newIssues = validateOrderRow(orderRow, pathVals);
      issues.push(...newIssues);
    });
  
    const issuesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ISSUES_SHEET_NAME);
    updateSheetFromObjects(issuesSheet, issues);
  }
  
  
  function validateOrderRow(orderRow, pathVals) {
    const issues = [];
    issues.addIssue = function (item, issue) {
      issues.push({
        'Item': item,
        'Issue': issue
      })
    }
    const path = pathVals.getPath(orderRow['Inventory ID'], orderRow['Path ID']);
    
    // Check to see if row meets MOQ
    const moq = path['MOQ'];
    if(orderRow['Quantity'] <= moq) {
      issues.addIssue(orderRow['Item'], Utilities.formatString(
        "Our MOQ on this item is %d. Let us know if we can increase the quantity you requested.", moq));
    }
  
    // Check to see if row meets Quantity Increment
    const quanInc = path['Quantity Increment'];
    if(orderRow['Quantity'] % quanInc !== 0) {
      issues.addIssue(orderRow['Item'], Utilities.formatString(
        "This item needs to be ordered in multiples of %d. Let us know if we can increase the quantity you requested.", quanInc));
    }
    return issues;
    
  }