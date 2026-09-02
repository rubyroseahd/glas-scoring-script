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
  return normalizedRegisteredVendors.indexOf(normalizedVendor) !== -1;
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

function parseVirtualSkuPrefixes(raw) {
  return safeStr(raw)
    .split(",")
    .map(v => safeStr(v).toUpperCase())
    .filter(Boolean);
}

function resolveFulfillmentType(sku, ingestedFulfillment, virtualSkuPrefixes) {
  const normalizedSku = safeStr(sku).toUpperCase();
  const normalizedIngested = safeStr(ingestedFulfillment).toUpperCase();
  const prefixes = Array.isArray(virtualSkuPrefixes) ? virtualSkuPrefixes : [];
  const hasVirtualPrefix = prefixes.some(prefix => normalizedSku.indexOf(safeStr(prefix).toUpperCase()) === 0);
  if (hasVirtualPrefix) return "WEBONLY";
  if (normalizedIngested === "WEBONLY") return "WEBONLY";
  return "SHARED";
}

function resetBiFeedSheet(sheet) {
  if (!sheet) return;
  if (typeof sheet.clearContents === "function") sheet.clearContents();
  resetSheetColumnGroups(sheet);
}

function safelyCollapseColumnGroup(sheet, columnIndex, depth) {
  try {
    if (sheet && typeof sheet.getColumnGroupDepth === "function" &&
        typeof sheet.collapseColumnGroup === "function" &&
        sheet.getColumnGroupDepth(columnIndex) >= depth) {
      sheet.collapseColumnGroup(columnIndex, depth);
    }
  } catch (err) {
    Logger.log(`[WARN] Skipping column group collapse at index ${columnIndex}: ${err.message}`);
  }
}

function resetSheetColumnGroups(sheet) {
  if (!sheet) return;
  try {
    if (typeof sheet.getMaxColumns !== "function" ||
        typeof sheet.getColumnGroupDepth !== "function" ||
        typeof sheet.removeColumnGroup !== "function") return;
    const maxCols = sheet.getMaxColumns();
    for (let col = maxCols; col >= 1; col--) {
      const depth = sheet.getColumnGroupDepth(col);
      for (let d = depth; d >= 1; d--) {
        sheet.removeColumnGroup(col);
      }
    }
  } catch (e) {
    Logger.log(`[WARN] Column group reset bypassed: ${e.message}`);
  }
}