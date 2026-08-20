// Requires: Apps Script editor → Services → Google Sheets API (Advanced Service)

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

    const lastRow = sourceSheet.getLastRow();
    const lastCol = Math.min(sourceSheet.getLastColumn(), 21); // cap at column U

    if (lastRow === 0) {
      Logger.log('No data found in source sheet.');
      return;
    }

    const data = sourceSheet.getRange(1, 1, lastRow, lastCol).getValues();
    Logger.log(`Fetched ${data.length} rows x ${lastCol} columns from source.`);

    writeToTarget(target1SpreadsheetId, target1TabName, data, 'Target 1');
    writeToTarget(target2SpreadsheetId, target2TabName, data, 'Target 2');

  } catch (e) {
    Logger.log('Error: ' + e.toString());
  }
}

/**
 * Uses Sheets REST API (Advanced Service) to clear and write data.
 * Avoids SpreadsheetApp.openById() which loads the full sheet model and times out on large sheets.
 */
function writeToTarget(spreadsheetId, tabName, data, label) {
  Logger.log(`Updating ${label}...`);

  const clearRange = `'${tabName}'!A:U`;
  const writeRange = `'${tabName}'!A1`;

  // Clear existing data
  Sheets.Spreadsheets.Values.clear({}, spreadsheetId, clearRange);

  // Write all rows in one API call with RAW input (no formula parsing overhead)
  Sheets.Spreadsheets.Values.update(
    { values: data, majorDimension: 'ROWS' },
    spreadsheetId,
    writeRange,
    { valueInputOption: 'RAW' }
  );

  Logger.log(`Successfully pasted ${data.length} rows into ${label}!`);
}
