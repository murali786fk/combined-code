function copyStoreDetailsDaily() {
  const sourceSpreadsheetId = '1JZG1EXD-eMCMZu7GN2vCreh-zxqMjRZGrKRJWgL1A88';
  const sourceTabName = 'Store Details';

  const target1SpreadsheetId = '1wFZnIttIVCEUUhKQoaNcK3d2j6I2_zDFQrWFk_ABydc';
  const target1TabName = 'Store Details';

  const target2SpreadsheetId = '1qU3DEy5hhHwyN0p9iokqT5K5j7GK3GM78ta_R4zK2nA';
  const target2TabName = 'Store_details';

  try {
    Logger.log('Fetching data from source sheet...');
    const sourceSheet = SpreadsheetApp.openById(sourceSpreadsheetId).getSheetByName(sourceTabName);

    // getLastRow() returns the last row with ANY content across all columns — reliable
    const lastRow = sourceSheet.getLastRow();
    const lastCol = Math.min(sourceSheet.getLastColumn(), 21); // cap at column U

    if (lastRow === 0) {
      Logger.log('No data found in source sheet.');
      return;
    }

    const data = sourceSheet.getRange(1, 1, lastRow, lastCol).getValues();
    Logger.log(`Fetched ${data.length} rows x ${lastCol} columns from source.`);

    updateTargetSheet(target1SpreadsheetId, target1TabName, data, 'Target 1');
    updateTargetSheet(target2SpreadsheetId, target2TabName, data, 'Target 2');

  } catch (e) {
    Logger.log('Error: ' + e.toString());
  }
}

function updateTargetSheet(spreadsheetId, tabName, data, label) {
  Logger.log(`Updating ${label}...`);
  const ss = SpreadsheetApp.openById(spreadsheetId);
  let sheet = ss.getSheetByName(tabName);

  if (!sheet) {
    sheet = ss.insertSheet(tabName);
  }

  const numRows = data.length;
  const numCols = data[0].length;
  const currentMaxRows = sheet.getMaxRows();

  sheet.clearContents(); // no flush here — batched with the write below

  if (currentMaxRows < numRows) {
    sheet.insertRowsAfter(currentMaxRows, numRows - currentMaxRows);
  }

  sheet.getRange(1, 1, numRows, numCols).setValues(data);
  SpreadsheetApp.flush(); // single sync at the end

  Logger.log(`Successfully pasted ${numRows} rows into ${label}!`);
}
