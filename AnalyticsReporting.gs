/**
 * MODULE 4: ANALYTICS REPORTING
 */

function generateAllReports() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dash = ss.getSheetByName(VDM_CONFIG.TABS.DASHBOARD);
    const data = dash.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);
    if (rows.length === 0) throw new Error("No data found in Dashboard Matrix.");
    
    // Load Shopify Memory Map for Handle and Vendor lookups
    const shopifyRaw = ss.getSheetByName(VDM_CONFIG.TABS.RAW_SHOPIFY).getDataRange().getValues();
    const shopifyMap = new Map(shopifyRaw.slice(1).map(r => {
      // Map SKU Anchor to Handle (index 2) and Vendor (index 7) based on ingestShopify logic
      return [r[0], { handle: r[2], vendor: r[7] }];
    }));

    const idx = {};
    headers.forEach((h, i) => idx[h] = i);

    generateSummaryTab(rows, idx, shopifyMap);
    generateSyncAudit(rows, idx, shopifyMap);
    generateMasterLedger(rows, idx, shopifyMap);
    generateSupplierScorecard(rows, idx, shopifyMap);
    logElasticitySnapshot(rows, idx);
  } catch (e) {
    logError("Reporting", e);
  }
}

function generateSummaryTab(rows, idx, shopifyMap) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(VDM_CONFIG.TABS.SUMMARY);
  sheet.clear().clearFormats();

  // Block 1: Migration Matrix
  const matrix = {};
  const tiers = ["Top Hero", "Signature Hero", "Proven Performer", "Accelerator", "Clearance", "New Launch", "Archive"];
  
  rows.forEach(r => {
    const from = r[idx["Current Equivalent Storefront Tier"]];
    const to = r[idx["Target Strategic Tier"]];
    if (!from || !to) return;

    const cost = parseFloat(r[idx["Resolved Cost Base"]]) || 0;
    // Extract clean name (remove % Off text) for the key
    const cleanTo = to.split(" (")[0];
    const key = `${from} -> ${cleanTo}`;

    if (!matrix[key]) matrix[key] = { count: 0, val: 0 };
    matrix[key].count++;
    matrix[key].val += cost;
  });

  const out = [["Migration Path", "SKU Count", "Invoiced Cost Value"]];
  Object.keys(matrix).forEach(k => out.push([k, matrix[k].count, matrix[k].val]));
  sheet.getRange(1, 1, out.length, 3).setValues(out);
  applyHeaderStyle(sheet.getRange(1, 1, 1, 3));

  // Block 2: House Brands
  const houseRows = rows.filter(r => {
    const sku = r[idx["SKU Anchor Key"]];
    const vendorName = (shopifyMap.get(sku)?.vendor || "").toUpperCase();
    return VDM_CONFIG.HOUSE_BRANDS.some(hb => vendorName.includes(hb.toUpperCase()));
  });
  // Simplified example logic for GLAS block
  sheet.getRange(out.length + 3, 1).setValue("House Brand Analytics").setFontWeight("bold");
}

function generateSyncAudit(rows, idx, shopifyMap) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(VDM_CONFIG.TABS.SYNC_AUDIT);
  sheet.clear().clearFormats();

  const headers = ["SKU", "Handle", "Action", "Final Tier", "Final Discount", "Old Variant Price", "Old Compare At Price", "Base Price Used", "New Variant Price", "New Compare At Price", "Note"];

  const syncRows = rows.map(r => {
    const sku = r[idx["SKU Anchor Key"]];
    const mkdn = parseFloat(r[idx["VDM Markdown Depth %"]]) || 0;
    return [
      sku,
      shopifyMap.get(sku)?.handle || "",
      r[idx["Pricing Migration Status"]],
      r[idx["Target Strategic Tier"]],
      mkdn,
      r[idx["Live Storefront Price"]],
      r[idx["Live Compare MSRP"]],
      r[idx["Live Compare MSRP"]],
      r[idx["New Proposed Storefront Price"]],
      mkdn === 0 ? "" : r[idx["Live Compare MSRP"]],
      ""
    ];
  });

  sheet.getRange(1, 1, 1, 11).setValues([headers]);
  applyHeaderStyle(sheet.getRange(1, 1, 1, 11));
  if (syncRows.length > 0) sheet.getRange(2, 1, syncRows.length, 11).setValues(syncRows);
}

function generateMasterLedger(rows, idx, shopifyMap) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(VDM_CONFIG.TABS.MASTER_LEDGER);
  sheet.clear().clearFormats();

  const headers = ["SKU", "Handle", "Gatekeeper Status", "Final Tier", "Action", "Old Variant Price", "Old Compare At Price", "Base Price Used", "Final Discount %", "New Variant Price", "New Compare At Price", "Simulated Checkout Net Price", "Resolved Cost Base", "Final Stacked Margin %", "Guardrail Alert"];
  
  const ledgerRows = rows.map(r => {
    const sku = r[idx["SKU Anchor Key"]];
    const mkdn = parseFloat(r[idx["VDM Markdown Depth %"]]) || 0;
    return [
      sku,
      shopifyMap.get(sku)?.handle || "",
      r[idx["Gatekeeper Status"]],
      r[idx["Target Strategic Tier"]],
      r[idx["Pricing Migration Status"]],
      r[idx["Live Storefront Price"]],
      r[idx["Live Compare MSRP"]],
      r[idx["Live Compare MSRP"]],
      mkdn,
      r[idx["New Proposed Storefront Price"]],
      mkdn === 0 ? "" : r[idx["Live Compare MSRP"]],
      r[idx["Simulated Checkout Net Price"]],
      r[idx["Resolved Cost Base"]],
      r[idx["Final Simulated Stacked Margin %"]],
      r[idx["Profit Guardrail Status Alert"]]
    ];
  });

  sheet.getRange(1, 1, 1, 15).setValues([headers]);
  applyHeaderStyle(sheet.getRange(1, 1, 1, 15));
  if (ledgerRows.length > 0) sheet.getRange(2, 1, ledgerRows.length, 15).setValues(ledgerRows);
}

function generateSupplierScorecard(rows, idx, shopifyMap) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(VDM_CONFIG.TABS.SCORECARD);
  sheet.clear().clearFormats();

  const vendorTotals = {};
  rows.forEach(r => {
    const sku = r[idx["SKU Anchor Key"]];
    const vendor = shopifyMap.get(sku)?.vendor || "Unknown Vendor";
    const stockVal = (parseFloat(r[idx["Total On-Hand Warehouse Stock"]]) || 0) * (parseFloat(r[idx["Resolved Cost Base"]]) || 0);
    const units90 = parseFloat(r[idx["Retail Velocity Score Component"]]) || 0; // Using score as proxy or pull raw

    if (!vendorTotals[vendor]) vendorTotals[vendor] = { skus: 0, stockValue: 0, sales90: 0 };
    vendorTotals[vendor].skus++;
    vendorTotals[vendor].stockValue += stockVal;
    vendorTotals[vendor].sales90 += units90;
  });

  const out = [["Vendor/Brand", "Active SKU Count", "Total Warehouse Capital Value", "90D Velocity (Units)"]];
  Object.keys(vendorTotals).forEach(v => {
    out.push([v, vendorTotals[v].skus, vendorTotals[v].stockValue, vendorTotals[v].sales90]);
  });

  sheet.getRange(1, 1, out.length, 4).setValues(out);
  applyHeaderStyle(sheet.getRange(1, 1, 1, 4));
  sheet.getRange(2, 3, out.length - 1, 1).setNumberFormat("$#,##0.00");
}

function logElasticitySnapshot(rows, idx) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VDM_CONFIG.TABS.ELASTICITY);
  const date = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd");
  
  const snapshot = rows.map(r => [
    date,
    r[idx["SKU Anchor Key"]],
    r[idx["VDM Markdown Depth %"]],
    r[idx["Simulated Checkout Net Price"]],
    r[idx["Retail Velocity Score Component"]]
  ]);
  
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Snapshot Date", "SKU", "Markdown Depth", "Price", "Velocity"]);
    applyHeaderStyle(sheet.getRange(1, 1, 1, 5));
  }
  sheet.getRange(sheet.getLastRow() + 1, 1, snapshot.length, 5).setValues(snapshot);
}

function executeFlexibleRefreshProcess() {
  const ui = SpreadsheetApp.getUi();
  try {
    ui.showModelessDialog(HtmlService.createHtmlOutput("<b>Processing VDM Refresh...</b>"), "System Status");
    runDataIngestion();
    executeDashboardRefresh();
    generateAllReports();
    ui.alert("VDM Refresh Complete.");
  } catch (e) {
    ui.alert("Process Failed: " + e.message);
  }
}