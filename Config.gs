/**
 * MODULE 1: GLOBAL ENVIRONMENT & SYSTEM VARIABLES MATRIX
 * System Version: 2.5.0 (Enterprise Production Locked)
 */

const VDM_CONFIG = {
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
    SETTINGS: "[01] Settings & Registries",
    DASHBOARD: "[02] Dashboard Matrix", 
    SUMMARY: "[03] Tier Summary & Distribution Panel", 
    SCORECARD: "[01] Supplier Scorecard & Capital Velocity", 
    ELASTICITY: "[04] Pricing Elasticity & Lift Analytics Ledger", 
    SYNC_AUDIT: "[07] Storefront Update & Sync Audit",
    MASTER_LEDGER: "[09] Master Pricing & Margin Ledger",
    
    // Hidden Ingestion
    RAW_SHOPIFY: "_raw_shopify",
    RAW_EEI_USA: "_raw_eei_usa",
    RAW_EEI_WEB: "_raw_eei_web",
    RAW_SALES: "_raw_sales",
    RAW_COST: "_raw_cost",
    MASTER_COST: "_resolved_cost_base"
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

  // Header String Registry for dynamic matching
  HEADERS: {
    SHOPIFY: ["VARIANT SKU", "HANDLE", "STATUS", "VARIANT PRICE", "VARIANT COMPARE AT PRICE", "VENDOR", "VARIANT INVENTORY QTY", "COST PER ITEM"],
    USA_WAREHOUSE: ["ITEM CODE", "EEI USA WAREHOUSE ON HAND STOCK", "SALES PAST 30 DAYS"],
    WEB_WAREHOUSE: ["ITEM CODE", "EEI WEB WAREHOUSE ON HAND STOCK", "SALES PAST 30 DAYS"],
    RETAIL_VELOCITY: ["PRODUCT VARIANT SKU", "NET ITEMS SOLD"],
    COST_WATERFALL: ["SKU", "EEI LAST PURCHASE PRICE", "GLAS COSTING", "COTR LAST PURCHASE PRICE"]
  }
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