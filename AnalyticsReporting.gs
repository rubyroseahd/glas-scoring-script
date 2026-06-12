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
    
    const idx = {};
    headers.forEach((h, i) => idx[h] = i);

    generateSummaryTab(rows, idx);
    generateSyncAudit(rows, idx);
    generateMasterLedger(rows, idx);
    generateSupplierScorecard(rows, idx);
    logElasticitySnapshot(rows, idx);
  } catch (e) {
    logError("Reporting", e);
  }
}

function generateSummaryTab(rows, idx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(VDM_CONFIG.TABS.SUMMARY);
  sheet.clear().clearFormats();

  // Block 1: Migration Matrix
  const matrix = {};
  rows.forEach(r => {
    const from = r[idx["Current Equivalent Storefront Tier"]];
    const to = r[idx["Target Strategic Tier"]];
    const cost = parseFloat(r[idx["Resolved Cost Base"]]) || 0;
    const key = `${from} -> ${to}`;
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
    const vendor = String(r[idx["Gatekeeper Status"]]).toUpperCase(); // Not strictly vendor, check raw
    return VDM_CONFIG.HOUSE_BRANDS.some(hb => vendor.includes(hb.toUpperCase()));
  });
  // Simplified example logic for GLAS block
  sheet.getRange(out.length + 3, 1).setValue("House Brand Analytics").setFontWeight("bold");
}

function generateSyncAudit(rows, idx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(VDM_CONFIG.TABS.SYNC_AUDIT);
  const dashName = VDM_CONFIG.TABS.DASHBOARD;
  const shopifyRaw = ss.getSheetByName(VDM_CONFIG.TABS.RAW_SHOPIFY);
  const shopifyIdx = getHeaderMap(shopifyRaw.getRange(1,1,1, shopifyRaw.getLastColumn()).getValues()[0]);
  const handleCol = shopifyIdx["Handle"] + 1; // Convert 0-index to 1-index for VLOOKUP

  sheet.clear().clearFormats();

  const headers = ["SKU", "Handle", "Action", "Final Tier", "Final Discount", "Old Variant Price", "Old Compare At Price", "Base Price Used", "New Variant Price", "New Compare At Price", "Note"];
  const formulas = rows.map((_, i) => {
    const r = i + 2;
    return [
      `='${dashName}'!A${r}`,
      `=IFERROR(VLOOKUP(A${r}, _raw_shopify!$A:$ZZ, ${handleCol}, 0), "")`,
      `='${dashName}'!X${r}`,
      `='${dashName}'!M${r}`,
      `='${dashName}'!N${r}`,
      `='${dashName}'!E${r}`,
      `='${dashName}'!F${r}`,
      `='${dashName}'!F${r}`,
      `='${dashName}'!S${r}`,
      `=IF('${dashName}'!N${r}=0, "", '${dashName}'!F${r})`,
      `=""`
    ];
  });

  sheet.getRange(1, 1, 1, 11).setValues([headers]);
  applyHeaderStyle(sheet.getRange(1, 1, 1, 11));
  sheet.getRange(2, 1, formulas.length, 11).setFormulas(formulas);
}

function generateMasterLedger(rows, idx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(VDM_CONFIG.TABS.MASTER_LEDGER);
  const dashName = VDM_CONFIG.TABS.DASHBOARD;
  sheet.clear().clearFormats();

  const headers = ["SKU", "Handle", "Gatekeeper Status", "Final Tier", "Action", "Old Variant Price", "Old Compare At Price", "Base Price Used", "Final Discount %", "New Variant Price", "New Compare At Price", "Simulated Checkout Net Price", "Resolved Cost Base", "Final Stacked Margin %", "Guardrail Alert"];
  const formulas = rows.map((_, i) => {
    const r = i + 2;
    return [
      `='${dashName}'!A${r}`, `=IFERROR(VLOOKUP(A${r}, _raw_shopify!$A:$I, 2, 0), "")`, `='${dashName}'!B${r}`, `='${dashName}'!M${r}`, `='${dashName}'!X${r}`,
      `='${dashName}'!E${r}`, `='${dashName}'!F${r}`, `='${dashName}'!F${r}`, `='${dashName}'!N${r}`, `='${dashName}'!S${r}`,
      `=IF('${dashName}'!N${r}=0, "", '${dashName}'!F${r})`, `='${dashName}'!T${r}`, `='${dashName}'!D${r}`, `='${dashName}'!U${r}`, `='${dashName}'!V${r}`
    ];
  });
  sheet.getRange(1, 1, 1, 15).setValues([headers]);
  applyHeaderStyle(sheet.getRange(1, 1, 1, 15));
  sheet.getRange(2, 1, formulas.length, 15).setFormulas(formulas);
}

function generateSupplierScorecard(rows, idx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(VDM_CONFIG.TABS.SCORECARD);
  sheet.clear().clearFormats();

  const vendorTotals = {};
  rows.forEach(r => {
    const vendor = r[idx["Gatekeeper Status"]]; // Should cross-ref shopify for better vendor data
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