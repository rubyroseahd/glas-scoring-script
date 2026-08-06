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

function escapeRegExp(value) {
  return safeStr(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function isMapVendorMatch(vendor, registeredVendors) {
  const normalizedVendor = safeStr(vendor).toUpperCase();
  if (!normalizedVendor) return false;

  const normalizedRegisteredVendors = (registeredVendors || []).map(name => safeStr(name).toUpperCase()).filter(Boolean);
  if (normalizedRegisteredVendors.indexOf(normalizedVendor) !== -1) return true;

  return normalizedRegisteredVendors.some(name => {
    if (name.length < 4) return false;
    return new RegExp("(^|\\W)" + escapeRegExp(name) + "($|\\W)").test(normalizedVendor);
  });
}

function getPercentileRankInc(sortedArr, value) {
  const n = sortedArr.length;
  if (n === 0) return 0;
  if (n === 1) return 1;
  let below = 0;
  let above = 0;
  for (let i = 0; i < n; i++) {
    if (sortedArr[i] < value) below++;
    else if (sortedArr[i] > value) above++;
  }
  const rankLow = below / (n - 1);
  const rankHigh = (n - 1 - above) / (n - 1);
  return (rankLow + rankHigh) / 2;
}

function getStorefrontBracket(price, compareMsrp) {
  const cleanPrice = safeNum(price);
  const cleanCompare = safeNum(compareMsrp);

  if (!mathGuard(cleanPrice, cleanCompare) || cleanCompare === 0 || cleanCompare === cleanPrice) {
    return VDM_CONFIG.BRACKET_NAMES.HERO;
  }

  const depth = (cleanCompare - cleanPrice) / cleanCompare;
  if (depth <= 0) return VDM_CONFIG.BRACKET_NAMES.HERO;
  if (depth < 0.35) return VDM_CONFIG.BRACKET_NAMES.SIGNATURE;
  if (depth < 0.45) return VDM_CONFIG.BRACKET_NAMES.PROVEN;
  if (depth < 0.55) return VDM_CONFIG.BRACKET_NAMES.ACCELERATOR;
  return VDM_CONFIG.BRACKET_NAMES.CLEARANCE;
}