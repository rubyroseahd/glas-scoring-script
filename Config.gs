/**
 * MODULE 1: GLOBAL ENVIRONMENT & SYSTEM VARIABLES MATRIX
 * Principal Architect: Gemini Code Assist / Enterprise Solutions Architect
 */

const VDM_CONFIG = {
  FOLDER_ID: "1m1BoV4XOYoHSCu1QeOOddmcOdNLUGdlh",
  
  SOURCE_FILES: {
    SHOPIFY: "shopify_export_gt.csv",      // Live storefront catalog
    EEI_USA: "EEI USA Whse Stock Report.csv", // Bulk wholesale
    EEI_WEB: "EEI WEB Whse Stock Report.csv", // Web storefront allocation
    SALES: "Total sales by product.csv",   // 90-day velocity ledger
    COST: "Cost_Data.csv"                  // Procurement cost database
  },

  TABS: {
    // User Facing
    BRIEF: "[00] Executive Brief", 
    ACTION: "[00] Action Items & Sign-off", 
    CONTROL: "[01] Control Panel", 
    DASHBOARD: "[02] Dashboard Matrix", 
    SUMMARY: "[03] Tier Summary & Distribution Panel", 
    SCORECARD: "[01] Supplier Scorecard & Capital Velocity", 
    ELASTICITY: "[04] Pricing Elasticity & Lift Analytics Ledger", 
    AGING: "[05] Warehouse Aging & Alternative Liquidation Workspace", 
    COMPLIANCE: "[06] MAP Compliance & Marketplace Competitor Variance", 
    SYNC: "[07] Storefront Update & Sync Audit",
    
    // Hidden Ingestion
    RAW_SHOPIFY: "_raw_shopify",
    RAW_EEI_USA: "_raw_eei_usa",
    RAW_EEI_WEB: "_raw_eei_web",
    RAW_SALES: "_raw_sales",
    RAW_COST: "_raw_cost",
    MASTER_COST: "_resolved_cost_base"
  },

  DESIGN: {
    HEADER_BG: "#000000",           // Solid Deep Black
    HEADER_TEXT: "#FFFFFF",         // White Bold
    ALERT_BREACH_BG: "#FCE8E6",     // Soft Light-Red
    ALERT_BREACH_TEXT: "#A51D24",   // Dark-Red
    ALERT_GWP_BG: "#E8F0FE",        // Soft Pastel-Blue
    ALERT_GWP_TEXT: "#1A73E8",      // Muted Navy
    ALERT_LAUNCH_BG: "#E6F4EA",     // Soft Pastel-Green
    ALERT_LAUNCH_TEXT: "#137333"    // Forest-Green
  },

  HOUSE_BRANDS: ["Gläs", "glastoy", "GLASTOY"],
  MIN_PROFIT_GUARDRAIL: 0.20
};

/**
 * Global Error Logger
 */
function logError(module, error) {
  Logger.log(`[ERROR][${module}] ${error.toString()}`);
  console.error(`[${module}]`, error);
}