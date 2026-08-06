/**
 * MODULE 1: GLOBAL ENVIRONMENT & SYSTEM VARIABLES MATRIX
 * System Version: 3.0.0-PROD
 */

const VDM_CONFIG = {
  VERSION: "3.0.0-PROD",
  AFFILIATE_RATE_DEFAULT: 0.15,
  PROFIT_FLOOR_GUARDRAIL: 0.20,
  CLEARANCE_CAP_WARN: 0.20,
  HERO_POOL_MIN_WARN: 0.15,
  FOLDER_ID: "1m1BoV4XOYoHSCu1QeOOddmcOdNLUGdlh",
  
  SOURCE_FILES: {
    SHOPIFY: "shopify_export_gt.csv",
    EEI_USA: "EEI USA Whse Stock Report.csv",
    EEI_WEB: "EEI WEB Whse Stock Report.csv",
    SALES: "Total sales by product.csv",
    COST: "Cost_Data.csv"
  },

  TABS: {
    // User-Facing
    BRIEF: "[00] Executive Brief", 
    ACTION: "[00] Action Items & Sign-off",
    SETTINGS: "[01] Control Panel",
    BASELINE: "[01b] Baseline Snapshot",
    DASHBOARD: "[02] Dashboard Matrix", 
    SUMMARY: "[03] Tier Summary & Distribution Panel", 
    SCORECARD: "[01] Supplier Scorecard & Capital Velocity", 
    ELASTICITY: "[04] Pricing Elasticity & Lift Analytics Ledger", 
    SYNC_AUDIT: "[07] Storefront Update & Sync Audit",
    MASTER_LEDGER: "[09] Master Pricing & Margin Ledger",
    
    // Workbook ingestion tabs
    RAW_SHOPIFY: "shopify_export",
    RAW_EEI_USA: "eei_usa_whse",
    RAW_EEI_WEB: "eei_web_whse",
    RAW_SALES: "shopify_90day_sales",
    RAW_COST: "cost_ledger",
    MASTER_COST: "resolved_cost_ledger",
    BACKUP: "_backup_matrix_data"
  },

  DESIGN: {
    HEADER_BG: "#000000",
    HEADER_TEXT: "#FFFFFF",
    ALERT_BREACH_BG: "#FCE8E6",
    ALERT_BREACH_TEXT: "#A51D24",
    ALERT_GWP_BG: "#E8F0FE",
    ALERT_GWP_TEXT: "#1A73E8",
    ALERT_LAUNCH_BG: "#E6F4EA",
    ALERT_LAUNCH_TEXT: "#137333",
    PANEL_GLOBAL_BG: "#444444",
    PANEL_PROPRIETARY_BG: "#1C3A27"
  },

  HOUSE_BRANDS: ["Gläs", "glastoy", "GLASTOY"],
  BRACKET_NAMES: {
    HERO: "Full MSRP / Hero",
    SIGNATURE: "Light Promo / Signature",
    PROVEN: "Moderate Promo / Proven",
    ACCELERATOR: "Deep Promo / Accelerator",
    CLEARANCE: "Clearance / Archive"
  },

  // Header String Registry for dynamic matching
  HEADERS: {
    SHOPIFY: ["VARIANT SKU", "HANDLE", "STATUS", "VARIANT PRICE", "VARIANT COMPARE AT PRICE", "VENDOR", "TYPE", "PRODUCT TYPE", "COST PER ITEM"],
    USA_WAREHOUSE: ["ITEM CODE", "SKU", "EEI USA WAREHOUSE ON HAND STOCK", "QTY", "QUANTITY", "SALES PAST 30 DAYS"],
    WEB_WAREHOUSE: ["ITEM CODE", "SKU", "EEI WEB WAREHOUSE ON HAND STOCK", "QTY", "QUANTITY", "SALES PAST 30 DAYS"],
    RETAIL_VELOCITY: ["PRODUCT VARIANT SKU", "VARIANT SKU", "SKU", "NET QUANTITY", "NET ITEMS SOLD", "QTY", "QUANTITY"],
    COST_WATERFALL: ["SKU", "VARIANT SKU", "LAST PURCHASE PRICE REPORT EEI(UPDATE)", "GLAS COSTING", "LAST PURCHASE PRICE REPORT COTR(UPDATE)", "COST", "UNIT COST", "COST PER ITEM"]
  }
};

const GUARDRAIL_CODES = {
  SAFE: "SAFE",
  ERR_MISSING_COST: "ERR_MISSING_COST",
  ERR_NEGATIVE_MARGIN: "ERR_NEGATIVE_MARGIN",
  ERR_MARGIN_FLOOR_VIOLATOR: "ERR_MARGIN_FLOOR_VIOLATOR",
  WARN_B2B_HOLD: "WARN_B2B_HOLD"
};

const TIER_CODES = {
  GATEKEEPER: "TIER_00_GATEKEEPER",
  TOP_HERO: "TIER_10_TOP_HERO",
  SIG_HERO: "TIER_08_SIG_HERO",
  PROVEN: "TIER_06_PROVEN_PERFORMER",
  ACCELERATOR: "TIER_04_ACCELERATOR",
  CLEARANCE: "TIER_00_CLEARANCE"
};

const GATEKEEPER_CODES = {
  GWP: "GK_GWP",
  NEW_LAUNCH: "GK_NEW_LAUNCH",
  MAP: "GK_MAP",
  NONE: "NONE"
};

const QUEUE_CODES = {
  QUEUE_1A_COST: "Q1A_MISSING_COST",
  QUEUE_1A_MARGIN: "Q1A_NEGATIVE_MARGIN",
  QUEUE_1B_FLOOR: "Q1B_MARGIN_FLOOR",
  QUEUE_2_WEBONLY: "Q2_WEBONLY_REVIEW",
  QUEUE_3_CLEARANCE: "Q3_SHARED_CLEARANCE",
  NONE: "NONE"
};

function logError(module, error) {
  const msg = `[ERROR][${module}] ${error.stack || error}`;
  Logger.log(msg);
  console.error(msg);
}

/**
 * Standardized Header Index Mapping
 */
function getHeaderMap(headers) {
  if (!headers || !Array.isArray(headers)) return {};
  const map = {};
  headers.forEach((h, i) => {
    if (h !== null && h !== undefined && h.toString().trim() !== "") {
      map[h.toString().trim().toUpperCase()] = i;
    }
  });
  return map;
}

/**
 * Standardized header formatting across the reporting suite.
 */
function applyHeaderStyle(range) {
  range.setBackground(VDM_CONFIG.DESIGN.HEADER_BG)
       .setFontColor(VDM_CONFIG.DESIGN.HEADER_TEXT)
       .setFontWeight("bold")
       .setHorizontalAlignment("center");
}

/**
 * Type-safe string conversion
 */
function safeStr(val) {
  return val === null || val === undefined ? "" : String(val).trim();
}

/**
 * Type-safe number conversion with percentage/currency cleaning
 */
function safeNum(val) {
  if (val === null || val === undefined || val === "") return null;
  if (typeof val === "number") return isNaN(val) ? null : val;
  const str = String(val).trim();
  const n = parseFloat(str.replace(/[$,%\s]/g, "").replace(/,/g, ""));
  if (isNaN(n)) return null;
  return str.includes("%") ? n / 100 : n;
}

/**
 * Math validity check for division and comparisons
 */
function mathGuard(...values) {
  return values.every(v => typeof v === 'number' && v !== null && !isNaN(v));
}