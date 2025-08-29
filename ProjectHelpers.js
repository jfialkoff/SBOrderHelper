function getPathData() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PATHS_SHEET_NAME);
    const pathVals = getDataAsObjects(sheet);
    const result = {};
    pathVals.forEach(path => {
      result[path['Catalog Label']] = path;
    });
  
    result.getPath = function (invID, pathID) {
      return result[invID + '-P' + pathID]
    }
  
    return result;
  }
  