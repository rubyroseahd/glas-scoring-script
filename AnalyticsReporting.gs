/**
 * MODULE 4: ANALYTICS REPORTING
 */

function generateAllReports(dashboardState) {
  try {
    if (!dashboardState || !dashboardState.rows) throw new Error("Dashboard state missing for reporting.");
    const { rows, headers } = dashboardState;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shopifyMap = getShopifyMap();
    if (shopifyMap.size === 0) throw new Error("Shopify metadata missing. Please run ingestion.");

    const idx = getHeaderMap(headers); // Use the standardized helper for Dashboard columns

    generateSummaryTab(ss, rows, idx, shopifyMap);
    generateSyncAudit(ss, rows, idx, shopifyMap);
    generateMasterLedger(ss, rows, idx, shopifyMap);
    generateSupplierScorecard(ss, rows, idx, shopifyMap);
    logElasticitySnapshot(ss, rows, idx);
  } catch (e) {
    logError("Reporting", e);
  }
}

/**
 * Helper to load Shopify Memory Map for Handle and Vendor lookups
 */
function getShopifyMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(VDM_CONFIG.TABS.RAW_SHOPIFY);
  if (!rawSheet) return new Map();
  const data = rawSheet.getDataRange().getValues();
  const idx = getHeaderMap(data[0]);
  return new Map(data.slice(1).map(r => [r[0], { handle: r[idx["HANDLE"]], vendor: r[idx["VENDOR"]] }]));
}

function generateSummaryTab(ss, rows, idx, shopifyMap) {
  const sheet = ss.getSheetByName(VDM_CONFIG.TABS.SUMMARY);
  sheet.clear().clearFormats();
  
  const settingsData = ss.getSheetByName(VDM_CONFIG.TABS.SETTINGS).getDataRange().getValues();
  const globalAffiliateRate = (settingsData.length > 1 && safeNum(settingsData[1][4]) !== null) ? safeNum(settingsData[1][4]) : 0.15;

  // --- PANEL A: GLOBAL CATALOG COMPARATIVE DISTRIBUTION MATRIX ---
  const panelAHeaders = ["Strategic Pricing Bracket", "Current Shopify SKU Count", "Current Shopify Catalog %", "Optimized VDM SKU Count", "Optimized VDM Catalog %", "Net Allocation Weight Shift % (Difference)", "Base VDM Markdown Depth %", "Active Global Affiliate Rate Reference", "Blended Cumulative Stacked Discount %", "Risk Classification Profile"];
  const brackets = [
    { name: "Top Hero Bracket", mkdn: 0.00, risk: "None-Low", shopCheck: (m) => m === 0, vdmMatch: "Top Hero" },
    { name: "Signature Hero Bracket", mkdn: 0.30, risk: "Low-Med", shopCheck: (m) => m > 0 && m <= 0.35, vdmMatch: "Signature Hero" },
    { name: "Proven Performer Bracket", mkdn: 0.40, risk: "Med", shopCheck: (m) => m > 0.35 && m <= 0.45, vdmMatch: "Proven Performer" },
    { name: "Accelerator Bracket", mkdn: 0.50, risk: "Med-High", shopCheck: (m) => m > 0.45 && m <= 0.55, vdmMatch: "Accelerator" },
    { name: "Clearance/Archive Bracket", mkdn: 0.65, risk: "High", shopCheck: (m) => m > 0.55, vdmMatch: "Clearance/Archive" },
    { name: "New Launch Bracket", mkdn: 0.00, risk: "None", shopCheck: (m) => false, vdmMatch: "New Launch" },
    { name: "B2B Protection Hold Bracket", mkdn: 0.00, risk: "None", shopCheck: (m) => false, vdmMatch: "B2B Protection Hold" }
  ];

  const totalRows = rows.length;
  let panelAData = brackets.map(b => {
    const shopCount = rows.filter(r => b.shopCheck(safeNum(r[idx["ACTIVE STOREFRONT MARKDOWN DEPTH %"]]) || 0)).length;
    const vdmRows = rows.filter(r => r[idx["TARGET STRATEGIC TIER"]] && r[idx["TARGET STRATEGIC TIER"]].startsWith(b.vdmMatch));
    const vdmCount = vdmRows.length;

    const shopPct = shopCount / totalRows;
    const vdmPct = vdmCount / totalRows;
    const diff = vdmPct - shopPct;
    const stacked = 1 - ((1 - b.mkdn) * (1 - globalAffiliateRate));

    return [b.name, shopCount, shopPct, vdmCount, vdmPct, diff, b.mkdn, globalAffiliateRate, stacked, b.risk];
  });

  // Add Total Row for Panel A
  const panelATotals = ["Total Catalog Reconciliation", 
    panelAData.reduce((s, r) => s + r[1], 0), panelAData.reduce((s, r) => s + r[2], 0),
    panelAData.reduce((s, r) => s + r[3], 0), panelAData.reduce((s, r) => s + r[4], 0),
    panelAData.reduce((s, r) => s + r[5], 0), "", "", "", ""];
  panelAData.push(panelATotals);

  const panelAWidth = panelAHeaders.length;
  sheet.getRange(1, 1).setValue("GLOBAL CATALOG ALLOCATION SUMMARY MATRIX").setFontSize(14).setFontWeight("bold").setBackground(VDM_CONFIG.DESIGN.PANEL_GLOBAL_BG).setFontColor("#FFFFFF");
  sheet.getRange(2, 1, 1, panelAWidth).setValues([panelAHeaders]);
  applyHeaderStyle(sheet.getRange(2, 1, 1, panelAWidth));
  sheet.getRange(3, 1, panelAData.length, panelAWidth).setValues(panelAData);

  [3, 5, 6, 7, 8, 9].forEach(col => {
    sheet.getRange(3, col, panelAData.length, 1).setNumberFormat("0.00%");
  });

  // --- PANEL B: GLÄS & GLASTOY PROPRIETARY CATALOG DELTA PANEL ---
  const houseRows = rows.filter(r => {
    const sku = r[idx["SKU ANCHOR KEY"]];
    const vendorName = (shopifyMap.get(sku)?.vendor || "").toUpperCase();
    return VDM_CONFIG.HOUSE_BRANDS.some(hb => vendorName.includes(hb.toUpperCase()));
  });
  const totalHouseRows = houseRows.length;

  const panelBHeaders = ["Proprietary Strategic Bracket", "GLÄS Current Shopify SKU Count", "GLÄS Current Catalog %", "GLÄS Optimized VDM SKU Count", "GLÄS Optimized VDM Catalog %", "GLÄS Net Weight Shift % (Difference)", "VDM Base Discount %", "Final Stacked Checkout Discount"];
  
  let panelBData = brackets.map(b => {
    const shopCount = houseRows.filter(r => b.shopCheck(safeNum(r[idx["ACTIVE STOREFRONT MARKDOWN DEPTH %"]]) ?? 0)).length;
    const vdmCount = houseRows.filter(r => r[idx["TARGET STRATEGIC TIER"]].startsWith(b.vdmMatch)).length;
    
    const shopPct = totalHouseRows > 0 ? shopCount / totalHouseRows : 0;
    const vdmPct = totalHouseRows > 0 ? vdmCount / totalHouseRows : 0;
    const diff = vdmPct - shopPct;
    const stacked = 1 - ((1 - b.mkdn) * (1 - globalAffiliateRate));

    return [b.name, shopCount, shopPct, vdmCount, vdmPct, diff, b.mkdn, stacked];
  });

  // Add Total Row for Panel B
  const panelBTotals = ["Proprietary Reconciliation Total", 
    panelBData.reduce((s, r) => s + r[1], 0), panelBData.reduce((s, r) => s + r[2], 0),
    panelBData.reduce((s, r) => s + r[3], 0), panelBData.reduce((s, r) => s + r[4], 0),
    panelBData.reduce((s, r) => s + r[5], 0), "", ""];
  panelBData.push(panelBTotals);

  const panelBWidth = panelBHeaders.length;
  const startPanelB = 3 + panelAData.length + 4;
  sheet.getRange(startPanelB, 1).setValue("GLÄS & GLASTOY PROPRIETARY BRAND INSIGHTS PANEL").setFontSize(12).setFontWeight("bold").setBackground(VDM_CONFIG.DESIGN.PANEL_PROPRIETARY_BG).setFontColor("#FFFFFF");
  sheet.getRange(startPanelB + 1, 1, 1, panelBWidth).setValues([panelBHeaders]);
  applyHeaderStyle(sheet.getRange(startPanelB + 1, 1, 1, panelBWidth));
  sheet.getRange(startPanelB + 2, 1, panelBData.length, panelBWidth).setValues(panelBData);
  sheet.getRange(startPanelB + 1 + panelBData.length, 1, 1, 8).setFontWeight("bold");

  [3, 5, 6, 7, 8].forEach(col => {
    sheet.getRange(startPanelB + 2, col, panelBData.length, 1).setNumberFormat("0.00%");
  });

  // --- PANEL C: CHANNEL CLASS VERIFICATION BLOCK ---
  const startPanelC = startPanelB + 2 + panelBData.length + 3;
  const sharedTotal = rows.filter(r => r[idx["FULFILLMENT TAG"]] === "SHARED" && r[idx["PRICING MIGRATION STATUS"]] !== "⚠️ HOLD: B2B Volume Stable").length;
  const b2bTotal = rows.filter(r => r[idx["PRICING MIGRATION STATUS"]] === "⚠️ HOLD: B2B Volume Stable").length;
  const webTotal = rows.filter(r => r[idx["FULFILLMENT TAG"]] === "WEBONLY").length;

  const panelCData = [
    ["Channel Classification Layer", "Total Active Catalog SKUs Count"],
    ["SHARED Physical Layer", sharedTotal],
    ["B2BONLY Reserve Layer", b2bTotal],
    ["WEBONLY Virtual Layer", webTotal],
    ["Grand Total Catalog Reconciliation", rows.length]
  ];
  sheet.getRange(startPanelC, 1, 5, 2).setValues(panelCData);
  sheet.getRange(startPanelC, 1, 1, 2).setFontWeight("bold").setBackground("#EEEEEE");
  sheet.getRange(startPanelC + 4, 1, 1, 2).setFontWeight("bold").setBorder(true, null, null, null, null, null);
}

function generateSyncAudit(ss, rows, idx, shopifyMap) {
  const sheet = ss.getSheetByName(VDM_CONFIG.TABS.SYNC_AUDIT);
  sheet.clear().clearFormats();

  // Revised Schema Layer 3
  const headers = [
    "SKU Anchor Key", "Handle", "Action Required", "Optimized VDM Strategic Tier", "Optimized VDM Markdown %",
    "Old Live Variant Price", "New Proposed Variant Price", "Old Live Compare At Price", 
    "New Proposed Compare At Price", "Calculated Base Price Used", "Operational Guardrail Note"
  ];

  const syncRows = rows.map(r => {
    const sku = r[idx["SKU ANCHOR KEY"]];
    const mkdn = safeNum(r[idx["VDM MARKDOWN DEPTH %"]]) || 0;
    const status = r[idx["PRICING MIGRATION STATUS"]];
    const action = (status === "✓ Price Hold" || status === "⚠️ HOLD: B2B Volume Stable") ? "NO CHANGE" : "UPDATED";
    
    return [
      sku, shopifyMap.get(sku)?.handle || "", action,
      r[idx["TARGET STRATEGIC TIER"]], mkdn,
      r[idx["LIVE STOREFRONT PRICE"]], r[idx["NEW PROPOSED STOREFRONT PRICE"]],
      r[idx["LIVE COMPARE MSRP"]], mkdn === 0 ? "" : r[idx["LIVE COMPARE MSRP"]],
      r[idx["LIVE COMPARE MSRP"]], r[idx["GATEKEEPER STATUS"]]
    ];
  });

  const width = headers.length;
  sheet.getRange(1, 1, 1, width).setValues([headers]);
  applyHeaderStyle(sheet.getRange(1, 1, 1, width));
  if (syncRows.length > 0) {
    const range = sheet.getRange(2, 1, syncRows.length, width);
    range.setValues(syncRows);
    sheet.getRange(2, 5, syncRows.length, 1).setNumberFormat("0.00%");
    [6, 7, 8, 9, 10].forEach(col => sheet.getRange(2, col, syncRows.length, 1).setNumberFormat("0.00"));
  }
}

function generateMasterLedger(ss, rows, idx, shopifyMap) {
  const sheet = ss.getSheetByName(VDM_CONFIG.TABS.MASTER_LEDGER) || ss.insertSheet(VDM_CONFIG.TABS.MASTER_LEDGER);
  sheet.clear().clearFormats();

  // Revised Schema Layer 4 - Financial Sequencing
  const headers = [
    "SKU Key", "Handle", "Fulfillment Classification", "Gatekeeper Control Status",
    "Pricing Migration Status", "Recommended Strategic Tier", "VDM Target Markdown Depth %",
    "Old Live Storefront Price", "New Proposed Storefront Price", "Calculated Face Value Price Shift ($)",
    "Resolved Procurement Cost Base", "Simulated Checkout Net Price", "Final Simulated Stacked Margin %",
    "Profit Guardrail Status Indicator", "Net Operational Margin Shift %"
  ];
  
  const ledgerRows = rows.map(r => {
    return [
      r[idx["SKU ANCHOR KEY"]], shopifyMap.get(r[idx["SKU ANCHOR KEY"]])?.handle || "",
      r[idx["FULFILLMENT TAG"]], r[idx["GATEKEEPER STATUS"]],
      r[idx["PRICING MIGRATION STATUS"]], r[idx["TARGET STRATEGIC TIER"]], r[idx["VDM MARKDOWN DEPTH %"]],
      r[idx["LIVE STOREFRONT PRICE"]], r[idx["NEW PROPOSED STOREFRONT PRICE"]], r[idx["RETAIL PRICE SHIFT ($)"]],
      r[idx["RESOLVED COST BASE"]], r[idx["SIMULATED CHECKOUT NET PRICE"]], r[idx["FINAL SIMULATED STACKED MARGIN %"]],
      r[idx["PROFIT GUARDRAIL STATUS ALERT"]], r[idx["NET MARGIN CHANGE %"]]
    ];
  });

  const width = headers.length;
  sheet.getRange(1, 1, 1, width).setValues([headers]);
  applyHeaderStyle(sheet.getRange(1, 1, 1, width));
  if (ledgerRows.length > 0) {
    const range = sheet.getRange(2, 1, ledgerRows.length, width);
    range.setValues(ledgerRows);

    // Separate dollar formatting from percentage margin fields
    [8, 9, 10, 11, 12].forEach(col => sheet.getRange(2, col, ledgerRows.length, 1).setNumberFormat("0.00"));
    [7, 13, 15].forEach(col => sheet.getRange(2, col, ledgerRows.length, 1).setNumberFormat("0.00%"));
  }
}

function generateSupplierScorecard(ss, rows, idx, shopifyMap) {
  const sheet = ss.getSheetByName(VDM_CONFIG.TABS.SCORECARD);
  sheet.clear().clearFormats();

  const vendorTotals = {};
  rows.forEach(r => {
    const sku = r[idx["SKU ANCHOR KEY"]];
    const vendor = shopifyMap.get(sku)?.vendor || "Unknown Vendor";
    const stockVal = (safeNum(r[idx["TOTAL ON-HAND WAREHOUSE STOCK"]]) || 0) * (safeNum(r[idx["RESOLVED COST BASE"]]) || 0);
    const units90 = safeNum(r[idx["RETAIL VELOCITY SCORE COMPONENT"]]) || 0; 

    if (!vendorTotals[vendor]) vendorTotals[vendor] = { skus: 0, stockValue: 0, sales90: 0 };
    vendorTotals[vendor].skus++;
    vendorTotals[vendor].stockValue += stockVal;
    vendorTotals[vendor].sales90 += units90;
  });

  const out = [["Vendor/Brand", "Active SKU Count", "Total Warehouse Capital Value", "90D Velocity (Units)"]];
  Object.keys(vendorTotals).forEach(v => {
    out.push([v, vendorTotals[v].skus, vendorTotals[v].stockValue, vendorTotals[v].sales90]);
  });

  if (out.length > 0) {
    sheet.getRange(1, 1, out.length, 4).setValues(out);
    applyHeaderStyle(sheet.getRange(1, 1, 1, 4));
  }
  if (out.length > 1) {
    sheet.getRange(2, 3, out.length - 1, 1).setNumberFormat("$#,##0.00");
  }
}

function logElasticitySnapshot(ss, rows, idx) {
  const sheet = ss.getSheetByName(VDM_CONFIG.TABS.ELASTICITY);
  const date = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
  
  const snapshot = rows.map(r => [
    date,
    r[idx["SKU ANCHOR KEY"]],
    r[idx["VDM MARKDOWN DEPTH %"]],
    r[idx["SIMULATED CHECKOUT NET PRICE"]],
    r[idx["RETAIL VELOCITY SCORE COMPONENT"]]
  ]);
  
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Snapshot Date", "SKU", "Markdown Depth", "Price", "Velocity"]);
    applyHeaderStyle(sheet.getRange(1, 1, 1, 5));
  }
  if (snapshot.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, snapshot.length, 5).setValues(snapshot);
  }
}

/**
 * State Recovery Pattern: Recovers dashboard data from the sheet if needed for modular reporting.
 */
function recoverDashboardState() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName(VDM_CONFIG.TABS.DASHBOARD);
  const data = dashSheet.getDataRange().getValues();
  if (data.length < 2) throw new Error("Dashboard data not found. Run full sync first.");
  return { headers: data[0], rows: data.slice(1) };
}

function executeFlexibleRefreshProcess() {
  const ui = SpreadsheetApp.getUi();
  try {
    ui.showModelessDialog(HtmlService.createHtmlOutput("<b>Executing Full VDM System Sync...</b>"), "System Status");
    runDataIngestion();
    const dashboardState = executeDashboardRefresh();
    generateAllReports(dashboardState);
    ui.alert("Full VDM System Sync Complete.");
  } catch (e) {
    ui.alert("Process Failed: " + e.message);
  }
}

/** 
 * GATE 1: INGESTION CONTROLS 
 */
function workflowIngestInventoryOnly() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const folder = DriveApp.getFolderById(VDM_CONFIG.FOLDER_ID);
    ingestEEI(folder, VDM_CONFIG.SOURCE_FILES.EEI_USA, VDM_CONFIG.TABS.RAW_EEI_USA, ss);
    ingestEEI(folder, VDM_CONFIG.SOURCE_FILES.EEI_WEB, VDM_CONFIG.TABS.RAW_EEI_WEB, ss);
    ui.alert("Inventory Snapshot Sync Complete.");
  } catch (e) { ui.alert("Sync Failed: " + e.message); }
}

function workflowIngestMetadataOnly() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const folder = DriveApp.getFolderById(VDM_CONFIG.FOLDER_ID);
    ingestSalesCSV(folder, ss);
    ingestGenericCSV(folder, VDM_CONFIG.SOURCE_FILES.COST, VDM_CONFIG.TABS.RAW_COST, "SKU", ss);
    executeCostResolutionWaterfall();
    ui.alert("Commercial Metadata Refresh Complete.");
  } catch (e) { ui.alert("Refresh Failed: " + e.message); }
}

/** 
 * GATE 2: COMPUTE & SIMULATE CONTROLS 
 */
function workflowComputeOnly() {
  const ui = SpreadsheetApp.getUi();
  try {
    const dashboardState = executeDashboardRefresh();
    generateAllReports(dashboardState);
    ui.alert("Matrix Recalculation Complete.");
  } catch (e) { ui.alert("Calculation Failed: " + e.message); }
}

/** 
 * GATE 3: REPORTING & VIEW CONTROLS 
 */
function workflowReportSummaryOnly() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const { rows, headers } = recoverDashboardState();
    const idx = getHeaderMap(headers);
    const shopifyMap = getShopifyMap();
    generateSummaryTab(ss, rows, idx, shopifyMap);
    ui.alert("Executive Summary Updated.");
  } catch (e) { ui.alert("Update Failed: " + e.message); }
}

function workflowReportSyncOnly() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const { rows, headers } = recoverDashboardState();
    const idx = getHeaderMap(headers);
    const shopifyMap = getShopifyMap();
    generateSyncAudit(ss, rows, idx, shopifyMap);
    ui.alert("Shopify Sync Audit Generated.");
  } catch (e) { ui.alert("Generation Failed: " + e.message); }
}

function workflowReportLedgerOnly() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const { rows, headers } = recoverDashboardState();
    const idx = getHeaderMap(headers);
    const shopifyMap = getShopifyMap();
    generateMasterLedger(ss, rows, idx, shopifyMap);
    ui.alert("Master Ledger Refreshed.");
  } catch (e) { ui.alert("Refresh Failed: " + e.message); }
}