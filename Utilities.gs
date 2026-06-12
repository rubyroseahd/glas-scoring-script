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
 * Adheres to Section 4: Whitespace Trimming, Case Standardization, and Hidden Character Erasure.
 * @param {any} key
 * @return {string}
 */
function sanitizeKey(key) {
  if (key === null || key === undefined) return "";
  return String(key).trim().toUpperCase().replace(/[\r\n\t]+/g, "");
}

/**
 * Safely converts a value to a number, stripping currency symbols/commas.
 * Returns 0 instead of NaN for invalid data.
 * @param {any} value
 * @return {number}
 */
function safeNum(value) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === 'number') return isNaN(value) ? 0 : value;
  const cleanValue = value.toString().replace(/[$,\s]/g, "");
  const num = parseFloat(cleanValue);
  return isNaN(num) ? 0 : num;
}

/**
 * Safely sanitizes a string, returning an empty string for null/undefined.
 * @param {any} value
 * @return {string}
 */
function safeStr(value) {
  if (value === null || value === undefined) return "";
  return value.toString().trim();
}

/**
 * Creates a mapping object from a header row to avoid hardcoded indices.
 * Usage: const idx = getHeaderMap(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]);
 * @param {string[]} headers
 * @return {Object}
 */
function getHeaderMap(headers) {
  return headers.reduce((acc, header, index) => {
    if (header !== null && header !== undefined && header.toString().trim() !== "") { // Ensure header is not empty
      // Enforce uppercase keys for case-insensitive lookup robustness
      acc[header.toString().trim().toUpperCase()] = index;
    }
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