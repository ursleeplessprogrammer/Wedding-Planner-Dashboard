/**
 * Custom function to get the GID (sheet ID) of a sheet by its name.
 * Useful for creating dynamic hyperlinks to sheets.
 * @param {string} sheetName The name of the sheet.
 * @return {number} The GID of the sheet, or null if not found.
 * @customfunction
 */
function SHEETID(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  return sheet ? sheet.getSheetId() : null;
}
