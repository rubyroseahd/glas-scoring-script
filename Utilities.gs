/**
 * MODULE 0: UTILITY & HELPER FUNCTIONS
 * Shared resources for data sanitization, sheet management, and dynamic mapping.
 */

/**
 * Fetches a sheet by name or creates it with default styling if it doesn't exist.
 * @param {string} sheetName 
 * @param {boolean} isHidden
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateSheet(sheetName, isHidden = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (isHidden) sheet.hideSheet();
  }
  return sheet;
}

/**
 * Ensures keys (like SKUs) are consistent for lookups.
 * @param {any} key
 * @return {string}
 */
function sanitizeKey(key) {
  if (key === null || key === undefined) return "";
  return String(key).trim().toUpperCase();
}

/**
 * Creates a mapping object from a header row to avoid hardcoded indices.
 * Usage: const idx = getHeaderMap(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]);
 * @param {string[]} headers
 * @return {Object}
 */
function getHeaderMap(headers) {
  return headers.reduce((acc, header, index) => {
    const cleanHeader = header.toString().trim();
    acc[cleanHeader] = index;
    return acc;
  }, {});
}

/**
 * Standardizes header formatting across the reporting suite.
 * @param {GoogleAppsScript.Spreadsheet.Range} range
 */
function applyHeaderStyle(range) {
  range.setBackground(VDM_CONFIG.DESIGN.HEADER_BG)
       .setFontColor(VDM_CONFIG.DESIGN.HEADER_TEXT)
       .setFontWeight("bold")
       .setHorizontalAlignment("center");
}

/**
 * Formats a currency value for reporting.
 * @param {number} value
 * @return {string}
 */
function formatCurrency(value) {
  if (isNaN(value)) return "$0.00";
  return Utilities.formatString("$%.2f", value);
}