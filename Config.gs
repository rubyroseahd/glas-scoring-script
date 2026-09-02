/**
 * MODULE 1: GLOBAL CONFIGURATION
 */

const VDM_CONFIG = {
  VERSION: "3.1-LEAN",
  FOLDER_ID: "",
  PROFIT_FLOOR_GUARDRAIL: 0.20,

  SOURCE_FILES: {
    SHOPIFY: "shopify_export_gt.csv",
    EEI_USA: "EEI USA Whse Stock Report.csv",
    EEI_WEB: "EEI WEB Whse Stock Report.csv",
    SALES: "Total sales by product.csv",
    COST: "Cost_Data.csv"
  },

  TABS: {
    CONTROL: "[01] Control Panel",
    DASHBOARD: "[02] Dashboard Matrix",
    BI_FEED: "_BI_Data_Feed",
    ACTION_HUB: "[00] Action Items & Sign-off",
    TIER_SUMMARY: "[03] Tier Summary & Profitability Panel",
    SYNC_AUDIT: "[07] Storefront Update & Sync Audit",
    MASTER_LEDGER: "[09] Master Pricing & Margin Ledger",
    SHOPIFY_EXPORT: "shopify_export",
    EEI_USA: "eei_usa_whse",
    EEI_WEB: "eei_web_whse",
    SALES_90D: "shopify_90day_sales",
    COST_LEDGER: "cost_ledger",
    RESOLVED_COST: "resolved_cost_ledger",
    SHOPIFY_OUTPUT: "shopify_export_output",
    BACKUP_MATRIX_DATA: "_backup_matrix_data",

    // Backward-compatible aliases
    SETTINGS: "[01] Control Panel",
    ACTION: "[00] Action Items & Sign-off",
    ACTION_REORDER_HUB: "[00] Action Items & Reorder Hub",
    SUMMARY: "[03] Tier Summary & Profitability Panel",
    RAW_SHOPIFY: "shopify_export",
    RAW_EEI_USA: "eei_usa_whse",
    RAW_EEI_WEB: "eei_web_whse",
    RAW_SALES: "shopify_90day_sales",
    RAW_COST: "cost_ledger",
    MASTER_COST: "resolved_cost_ledger",
    SHOPIFY_EXPORT_OUTPUT: "shopify_export_output",
    BASELINE: "[01b] Baseline Snapshot"
  },

  FLOORS: {
    STACKED_MARGIN_MIN: 0.20,
    DRIFT_ALERT_RATIO: 0.20
  },

  DEFAULTS: {
    B2B_RESERVE_MIN_QTY: 500,
    AFFILIATE_RATE: 0.15,
    DOMESTIC_LEAD_DAYS: 30,
    IMPORT_LEAD_DAYS: 90,
    ENFORCE_OOS_REVERSION: true
  },

  BRACKET_NAMES: {
    HERO: "Full MSRP / Hero",
    SIGNATURE: "Light Promo / Signature",
    PROVEN: "Moderate Promo / Proven",
    ACCELERATOR: "Deep Promo / Accelerator",
    CLEARANCE: "Clearance / Archive"
  },

  HOUSE_BRANDS: ["GLÄS", "GLAS", "GLASTOY"]
};

const GUARDRAIL_CODES = {
  SAFE: "SAFE",
  WARN_B2B_HOLD: "WARN_B2B_HOLD",
  BLOCK_MARGIN_FLOOR: "BLOCK_MARGIN_FLOOR",

  // Backward-compatible aliases
  ERR_MISSING_COST: "GK_MISSING_COST",
  ERR_NEGATIVE_MARGIN: "GK_NEG_MARGIN",
  ERR_MARGIN_FLOOR_VIOLATOR: "BLOCK_MARGIN_FLOOR"
};

const GATEKEEPER_CODES = {
  NONE: "NONE",
  MISSING_COST: "GK_MISSING_COST",
  GWP: "GK_GWP",
  LAUNCH: "GK_LAUNCH",
  OOS: "GK_OOS",
  MAP: "GK_MAP",
  NEG_MARGIN: "GK_NEG_MARGIN",

  // Backward-compatible alias
  NEW_LAUNCH: "GK_LAUNCH"
};

const QUEUE_CODES = {
  NONE: "NONE",
  QUEUE_1A: "QUEUE_1A",
  QUEUE_1B: "QUEUE_1B",
  QUEUE_2: "QUEUE_2",
  QUEUE_3: "QUEUE_3",
  REORDER_ALERT: "REORDER_ALERT",

  // Backward-compatible aliases
  QUEUE_1A_COST: "QUEUE_1A",
  QUEUE_1A_MARGIN: "QUEUE_1A",
  QUEUE_1B_FLOOR: "QUEUE_1B",
  QUEUE_2_WEBONLY: "QUEUE_2",
  QUEUE_3_CLEARANCE: "QUEUE_3"
};

function logError(module, error) {
  const msg = `[ERROR][${module}] ${error && error.stack ? error.stack : error}`;
  Logger.log(msg);
}

function getHeaderMap(headers) {
  if (!Array.isArray(headers)) return {};
  const map = {};
  headers.forEach((h, i) => {
    const key = safeStr(h).toUpperCase();
    if (key) map[key] = i;
  });
  return map;
}

function applyHeaderStyle(range) {
  range.setFontWeight("bold").setHorizontalAlignment("center");
}

function safeStr(val) {
  return val === null || val === undefined ? "" : String(val).trim();
}

function safeNum(val) {
  if (val === null || val === undefined || val === "") return null;
  if (typeof val === "number") return isNaN(val) ? null : val;
  const str = String(val).trim();
  const n = parseFloat(str.replace(/[$,%\s]/g, "").replace(/,/g, ""));
  if (isNaN(n)) return null;
  return str.indexOf("%") !== -1 ? n / 100 : n;
}

function mathGuard() {
  return Array.prototype.slice.call(arguments).every(v => typeof v === "number" && !isNaN(v));
}

function resolveOperationalFolderId(propertyValue, legacyValue) {
  const resolved = (propertyValue && propertyValue.trim()) || (legacyValue && legacyValue.trim()) || "";
  if (!resolved) {
    throw new Error(
      "Missing VDM Drive folder configuration. " +
      "Set Script Property VDM_FOLDER_ID to the Drive folder ID, " +
      "or configure VDM_CONFIG.FOLDER_ID as a legacy fallback."
    );
  }
  return resolved;
}

function getOperationalFolderId() {
  const propertyValue = PropertiesService.getScriptProperties().getProperty("VDM_FOLDER_ID");
  return resolveOperationalFolderId(propertyValue, VDM_CONFIG.FOLDER_ID);
}

function loadControlPanelConfig(ss) {
  const sheet = ss.getSheetByName(VDM_CONFIG.TABS.CONTROL);
  const data = sheet ? sheet.getDataRange().getValues() : [];

  const skuSetFromColumn = colIndex => {
    const set = new Set();
    for (let i = 1; i < data.length; i++) {
      const key = safeStr(data[i][colIndex]).toUpperCase();
      if (key) set.add(key);
    }
    return set;
  };

  const vendorSetFromColumn = colIndex => {
    const set = new Set();
    for (let i = 1; i < data.length; i++) {
      const key = safeStr(data[i][colIndex]).toUpperCase();
      if (key) set.add(key);
    }
    return set;
  };

  const boolFromCell = value => {
    if (value === true || value === false) return value;
    const normalized = safeStr(value).toUpperCase();
    if (!normalized) return VDM_CONFIG.DEFAULTS.ENFORCE_OOS_REVERSION;
    return normalized === "TRUE" || normalized === "YES" || normalized === "1";
  };

  const b2bReserveMinRaw = data[1] ? safeNum(data[1][3]) : null;
  const affiliateRateRaw = data[1] ? safeNum(data[1][4]) : null;
  const virtualPrefixesRaw = data
    .slice(1)
    .map(r => safeStr(r[5]))
    .filter(Boolean)
    .join(",");
  const domesticLeadRaw = data[1] ? safeNum(data[1][6]) : null;
  const importLeadRaw = data[1] ? safeNum(data[1][7]) : null;
  const oosToggleRaw = data[1] ? data[1][8] : null;

  return {
    gwpSkus: skuSetFromColumn(0),
    newLaunchSkus: skuSetFromColumn(1),
    mapVendors: vendorSetFromColumn(2),
    b2bReserveMinQty: b2bReserveMinRaw !== null && b2bReserveMinRaw > 0
      ? Math.round(b2bReserveMinRaw)
      : VDM_CONFIG.DEFAULTS.B2B_RESERVE_MIN_QTY,
    affiliateRate: affiliateRateRaw !== null && affiliateRateRaw >= 0 && affiliateRateRaw < 1
      ? affiliateRateRaw
      : VDM_CONFIG.DEFAULTS.AFFILIATE_RATE,
    virtualSkuPrefixes: virtualPrefixesRaw
      ? parseVirtualSkuPrefixes(virtualPrefixesRaw)
      : [],
    defaultDomesticLeadTime: domesticLeadRaw !== null && domesticLeadRaw > 0
      ? Math.round(domesticLeadRaw)
      : VDM_CONFIG.DEFAULTS.DOMESTIC_LEAD_DAYS,
    defaultImportLeadTime: importLeadRaw !== null && importLeadRaw > 0
      ? Math.round(importLeadRaw)
      : VDM_CONFIG.DEFAULTS.IMPORT_LEAD_DAYS,
    enforceOosReversion: boolFromCell(oosToggleRaw)
  };
}
