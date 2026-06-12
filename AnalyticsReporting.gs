/**
 * MODULE 4: ANALYTICS REPORTING
 * Executive Summaries, Presentation Dashboards & Sync Audits
 */

function generateExecutiveSummary(dashRows, dIdx) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const summarySheet = getOrCreateSheet(VDM_CONFIG.TABS.SUMMARY);
    summarySheet.clear();

    const tiers = ['Top Hero', 'Signature Hero', 'Proven Performers', 'Accelerators', 'Clearance', 'New Launch', 'Excluded', 'Archive'];
    const summaryData = {};
    tiers.forEach(t => summaryData[t] = { skus: 0, shared: 0, b2b: 0, web: 0, depth: 0 });
    const channels = { 'SHARED': 0, 'B2BONLY': 0, 'WEBONLY': 0 };

    dashRows.forEach(r => {
      const tierStr = r[dIdx["Target Strategic Tier"]] || "";
      const gatekeeper = r[dIdx["Gatekeeper Status"]] || "";
      const fulfillment = r[dIdx["Fulfillment Tag"]] || "";

      let bucket = 'Archive';
      if (gatekeeper === "New Launch") bucket = 'New Launch';
      else if (gatekeeper === "⚠️ Active GWP Promo") bucket = 'Excluded';
      else if (tierStr.includes("Top Hero")) bucket = 'Top Hero';
      else if (tierStr.includes("Signature Hero")) bucket = 'Signature Hero';
      else if (tierStr.includes("Proven Performer")) bucket = 'Proven Performers';
      else if (tierStr.includes("Accelerator")) bucket = 'Accelerators';
      else if (tierStr.includes("Clearance")) bucket = 'Clearance';
      else if (tierStr.includes("Archive")) bucket = 'Archive';

      if (summaryData[bucket]) {
        summaryData[bucket].skus++;
        if (fulfillment === "SHARED") summaryData[bucket].shared++;
        else if (fulfillment === "B2BONLY") summaryData[bucket].b2b++;
        else if (fulfillment === "WEBONLY") summaryData[bucket].web++;
        summaryData[bucket].depth = r[dIdx["VDM Markdown Depth %"]] || 0;
      }
      if (channels[fulfillment] !== undefined) channels[fulfillment]++;
    });

    const controlSheet = ss.getSheetByName(VDM_CONFIG.TABS.CONTROL);
    const affiliateRate = controlSheet ? controlSheet.getRange("E2").getValue() : 0;
    const totalSkus = dashRows.length;

    const outputTable = [["Tier", "Total SKUs", "Shared", "B2B Only", "Web Only", "% Of Total", "Max Discount", "Affiliate ", "Final Discount", "", "Discount Tier", ""]];

    tiers.forEach(t => {
      const d = summaryData[t];
      const pctOfTotal = totalSkus > 0 ? d.skus / totalSkus : 0;
      outputTable.push([
        t, d.skus, d.shared, d.b2b, d.web, pctOfTotal, d.depth, affiliateRate, d.depth + affiliateRate, "", t, ""
      ]);
    });

    summarySheet.getRange(1, 1, outputTable.length, 12).setValues(outputTable);
    applyHeaderStyle(summarySheet.getRange(1, 1, 1, 12));
    summarySheet.getRange(2, 6, tiers.length, 1).setNumberFormat("0.00%");
    summarySheet.getRange(2, 7, tiers.length, 3).setNumberFormat("0%");

    const channelRow = outputTable.length + 3;
    summarySheet.getRange(channelRow, 1, 1, 2).setValues([["Channel Class", "Total SKUs"]]);
    applyHeaderStyle(summarySheet.getRange(channelRow, 1, 1, 2));
    summarySheet.getRange(channelRow + 1, 1, 3, 2).setValues([
      ["SHARED", channels["SHARED"]],
      ["B2BONLY", channels["B2BONLY"]],
      ["WEBONLY", channels["WEBONLY"]]
    ]);

    // These functions are now called from runFullRefreshCycle with the correct arguments
    // buildStorefrontSyncAudit(rows, dIdx);
    // logPricingElasticitySnapshot(rows, dIdx);
    // generateSupplierScorecard(rows, vendorMap);
    // generateWarehouseAgingReport(rows);
    // generateMAPComplianceReport(rows);

    console.log("Analytics Reporting Complete.");
  } catch (e) {
    logError("Analytics", e);
  }
}

function logPricingElasticitySnapshot(rows, dIdx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ledger = getOrCreateSheet(VDM_CONFIG.TABS.ELASTICITY);
  const dateStamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
  
  const logData = rows.map(r => [
    dateStamp,
    r[dIdx["SKU Anchor Key"]],
    r[dIdx["VDM Markdown Depth %"]],
    r[dIdx["Simulated Checkout Net Price"]], // Log checkout price for elasticity
    r[dIdx["Velocity Score Component"]]
  ]);
  // Fix 4: Corrected destructive overwriting sequence
  if (ledger.getLastRow() === 0) {
    ledger.getRange(1, 1, 1, 5).setValues([["Snapshot Date", "SKU", "Markdown Depth %", "Checkout Price", "Velocity Score"]]);
    applyHeaderStyle(ledger.getRange(1, 1, 1, 5));
  }
  if (logData.length > 0) { // Only write if there's data
    ledger.getRange(ledger.getLastRow() + 1, 1, logData.length, 5).setValues(logData);
  }
}

function generateSupplierScorecard(rows, vendorMap) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(VDM_CONFIG.TABS.SCORECARD);
  sheet.clear();
  
  const dIdx = getHeaderMap(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VDM_CONFIG.TABS.DASHBOARD).getDataRange().getValues()[0]);
  const supplierData = {};

  rows.forEach(r => {
    const sku = r[dIdx["SKU Anchor Key"]];
    const vendor = vendorMap[sku] || "Unknown";
    const costOfStock = (parseFloat(r[dIdx["Resolved Cost Base"]]) || 0) * (parseFloat(r[dIdx["Total On-Hand Warehouse Stock"]]) || 0);

    if (!supplierData[vendor]) {
      supplierData[vendor] = { skus: 0, totalValue: 0, velocitySum: 0 };
    }
    supplierData[vendor].skus++;
    supplierData[vendor].totalValue += costOfStock;
    supplierData[vendor].velocitySum += parseFloat(r[dIdx["Velocity Score Component"]]) || 0;
  });

  const output = [["Vendor / Brand", "Active SKU Count", "Total Invoiced Stock Value", "Avg Velocity Score"]];
  Object.keys(supplierData).forEach(v => {
    const s = supplierData[v];
    output.push([v, s.skus, s.totalValue, s.skus > 0 ? s.velocitySum / s.skus : 0]);
  });

  sheet.getRange(1, 1, output.length, 4).setValues(output);
  applyHeaderStyle(sheet.getRange(1, 1, 1, 4)); // Fix 3: Use applyHeaderStyle
  sheet.getRange(2, 3, output.length - 1, 1).setNumberFormat("$#,##0.00");
}

function generateWarehouseAgingReport(rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(VDM_CONFIG.TABS.AGING);
  sheet.clear();
  
  const dIdx = getHeaderMap(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VDM_CONFIG.TABS.DASHBOARD).getDataRange().getValues()[0]);
  
  // Filter for Clearance/Archive items with 0 Velocity
  const agingRows = rows.filter(r => 
    r[dIdx["Target Strategic Tier"]].includes("Clearance/Archive") && 
    parseFloat(r[dIdx["Velocity Score Component"]]) === 0
  ).map(r => [
    r[dIdx["SKU Anchor Key"]],
    r[dIdx["Total On-Hand Warehouse Stock"]],
    (parseFloat(r[dIdx["Resolved Cost Base"]]) || 0) * (parseFloat(r[dIdx["Total On-Hand Warehouse Stock"]]) || 0),
    "Alternative Disposal Recommended"
  ]);

  const headers = [["SKU", "Units On Hand", "Capital Tied Up", "Recommendation"]];
  sheet.getRange(1, 1, 1, 4).setValues(headers); // Fix 3: Use applyHeaderStyle
  applyHeaderStyle(sheet.getRange(1, 1, 1, 4));
  if (agingRows.length > 0) sheet.getRange(2, 1, agingRows.length, 4).setValues(agingRows);
}

function generateMAPComplianceReport(rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(VDM_CONFIG.TABS.COMPLIANCE);
  sheet.clear();

  const dIdx = getHeaderMap(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VDM_CONFIG.TABS.DASHBOARD).getDataRange().getValues()[0]);
  const mapRows = rows.filter(r => r[dIdx["Gatekeeper Status"]] === "3rd Party MAP")
                      .map(r => [r[dIdx["SKU Anchor Key"]], r[dIdx["Live Storefront Price"]], "Review Required"]);

  const headers = [["SKU", "Current Live Price", "Compliance Status"]];
  sheet.getRange(1, 1, 1, 3).setValues(headers); // Fix 3: Use applyHeaderStyle
  applyHeaderStyle(sheet.getRange(1, 1, 1, 3));
  if (mapRows.length > 0) sheet.getRange(2, 1, mapRows.length, 3).setValues(mapRows);
}

/**
 * NEW: Generates the Master Pricing & Margin Ledger presentation tab.
 */
function generateMasterLedgerTab(dashRows, dIdx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ledgerSheet = getOrCreateSheet(VDM_CONFIG.TABS.MASTER_LEDGER);
  
  if (dashRows.length === 0) return;

  ledgerSheet.clear();
  ledgerSheet.clearConditionalFormatRules();

  const dashName = VDM_CONFIG.TABS.DASHBOARD;
  const shopifyName = VDM_CONFIG.TABS.RAW_SHOPIFY;
  const shopifyRaw = ss.getSheetByName(shopifyName);
  const sIdx = getHeaderMap(shopifyRaw.getDataRange().getValues()[0]);
  const handleColIndex = sIdx["Handle"] + 1; // 1-indexed for VLOOKUP

  const formulaMatrix = dashRows.map((_, i) => {
    const r = i + 2; // Data starts on row 2
    return [
      `='${dashName}'!A${r}`,                                         // SKU
      `=VLOOKUP(A${r}, '${shopifyName}'!$A:$ZZ, ${handleColIndex}, 0)`, // Handle
      `='${dashName}'!B${r}`,                                         // Gatekeeper Status
      `='${dashName}'!M${r}`,                                         // Final Tier
      `=IF(ISNUMBER(SEARCH("✓", '${dashName}'!X${r})), "NO CHANGE", "UPDATED")`, // Action
      `='${dashName}'!E${r}`,                                         // Old Variant Price
      `='${dashName}'!F${r}`,                                         // Old Compare At Price
      `='${dashName}'!F${r}`,                                         // Base Price Used
      `='${dashName}'!N${r}`,                                         // Final Discount %
      `='${dashName}'!S${r}`,                                         // New Variant Price
      `=IF('${dashName}'!N${r}>0, '${dashName}'!F${r}, "")`,           // New Compare At Price
      `='${dashName}'!T${r}`,                                         // Simulated Checkout Net Price
      `='${dashName}'!D${r}`,                                         // Resolved Cost Base
      `='${dashName}'!U${r}`,                                         // Final Stacked Margin %
      `='${dashName}'!V${r}`                                          // Guardrail Alert
    ];
  });

  const headers = [["SKU", "Handle", "Gatekeeper Status", "Final Tier", "Action", "Old Variant Price", "Old Compare At Price", "Base Price Used", "Final Discount %", "New Variant Price", "New Compare At Price", "Simulated Checkout Net Price", "Resolved Cost Base", "Final Stacked Margin %", "Guardrail Alert"]];
  ledgerSheet.getRange(1, 1, 1, 15).setValues(headers);
  applyHeaderStyle(ledgerSheet.getRange(1, 1, 1, 15));
  ledgerSheet.getRange(2, 1, formulaMatrix.length, 15).setFormulas(formulaMatrix);

  // Formatting for Currency and Percentages
  ledgerSheet.getRange(2, 6, formulaMatrix.length, 3).setNumberFormat("$#,##0.00");
  ledgerSheet.getRange(2, 9, formulaMatrix.length, 1).setNumberFormat("0%");
  ledgerSheet.getRange(2, 10, formulaMatrix.length, 4).setNumberFormat("$#,##0.00");
  ledgerSheet.getRange(2, 14, formulaMatrix.length, 1).setNumberFormat("0%");

  // Visual Alerting Step: Conditional Formatting for Guardrail Alert (Row-wide)
  const range = ledgerSheet.getRange(2, 1, formulaMatrix.length, 15);
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$O2="❌ BLOCKED"')
    .setBackground(VDM_CONFIG.DESIGN.ALERT_BREACH_BG)
    .setFontColor(VDM_CONFIG.DESIGN.ALERT_BREACH_TEXT)
    .setBold(true)
    .setRanges([range])
    .build();
  ledgerSheet.setConditionalFormatRules([rule]);
}

function runFullRefreshCycle() {
  const ui = SpreadsheetApp.getUi();
  try {
    runDataIngestion();
    refreshDashboardMatrix();

    // --- NEW: Extract Dash Data once for all reporting modules ---
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dash = ss.getSheetByName(VDM_CONFIG.TABS.DASHBOARD);
    const dashData = dash.getDataRange().getValues();
    const dIdx = getHeaderMap(dashData[0]);
    const rows = dashData.slice(1);

    // Rebuild vendor map for the scorecard
    const shopifyRaw = ss.getSheetByName(VDM_CONFIG.TABS.RAW_SHOPIFY);
    const shopifyData = shopifyRaw.getDataRange().getValues();
    const sIdx = getHeaderMap(shopifyData[0]);
    const vendorMap = {};
    shopifyData.slice(1).forEach(r => {
      const itemSku = sanitizeKey(r[sIdx["SKU_ANCHOR"]]);
      if (itemSku) vendorMap[itemSku] = r[sIdx["Vendor"]];
    });

    generateExecutiveSummary(rows, dIdx);
    logPricingElasticitySnapshot(rows, dIdx);
    generateSupplierScorecard(rows, vendorMap);
    generateWarehouseAgingReport(rows);
    generateMAPComplianceReport(rows);
    generateMasterLedgerTab(rows, dIdx);

    ui.alert("VDM v2.2.2 Refresh Cycle Complete. 15-Minute Execution Cycle finished.");
  } catch (e) {
    logError("MainCycle", e);
    ui.alert("ERROR: " + e.message);
  }
}