/**
 * MODULE 4: ANALYTICS REPORTING
 * Executive Summaries, Presentation Dashboards & Sync Audits
 */

function generateExecutiveSummary(dashRows, dIdx) { // Fix 1: Accept arguments
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // const dash = ss.getSheetByName(VDM_CONFIG.TABS.DASHBOARD); // No longer needed
    // const dashData = dash.getDataRange().getValues(); // No longer needed
    // const dashHeaders = dashData[0]; // No longer needed
    const rows = dashRows; // Use passed argument

    // Block 1: Global Tier Migration Matrix
    const migrationMatrix = {};
    rows.forEach(r => {
      // Aligned with MatrixEngine.gs headers
      const from = r[dIdx["Current Equivalent Storefront Tier"]] || "N/A";
      const to = r[dIdx["Target Strategic Tier"]]; 
      const cost = parseFloat(r[dIdx["Resolved Cost Base"]]) || 0; 
      
      const key = `${from} -> ${to}`;
      if (!migrationMatrix[key]) migrationMatrix[key] = { count: 0, costImpact: 0 };
      migrationMatrix[key].count++;
      migrationMatrix[key].costImpact += cost;
    });

    const summarySheet = getOrCreateSheet(VDM_CONFIG.TABS.SUMMARY);
    summarySheet.clear();
    summarySheet.getRange("A1:C1").setValues([["Migration Coordinate", "SKU Frequency", "Invoiced Asset Cost"]])
      .setBackground("#000000").setFontColor("#FFFFFF").setFontWeight("bold");
    
    const migrationRows = Object.keys(migrationMatrix).map(k => [k, migrationMatrix[k].count, migrationMatrix[k].costImpact]);
    if (migrationRows.length > 0) {
      summarySheet.getRange(2, 1, migrationRows.length, 3).setValues(migrationRows);
    }

    // Block 2: Dedicated GLAS Summary Block
    const shopifyRaw = ss.getSheetByName(VDM_CONFIG.TABS.RAW_SHOPIFY);
    const shopifyData = shopifyRaw.getDataRange().getValues();
    const shopHeaders = shopifyData[0];
    const sIdx = getHeaderMap(shopHeaders);

    const vendorMap = {};
    // Fix 1: Corrected vendorMap creation - map SKU to Vendor Name
    shopifyData.slice(1).forEach(r => {
      const itemSku = sanitizeKey(r[sIdx["Variant SKU"]]);
      if (itemSku) {
        vendorMap[itemSku] = r[sIdx["Vendor"]];
      }
    });

    const glasSummary = {};
    rows.forEach(r => {
      const sku = r[dIdx["SKU Anchor Key"]];
      const vendor = (vendorMap[sku] || "").toLowerCase();
      const isHouseBrand = VDM_CONFIG.HOUSE_BRANDS.some(hb => vendor.includes(hb.toLowerCase()));

      if (isHouseBrand) {
        const tier = r[dIdx["Target Strategic Tier"]];
        if (!glasSummary[tier]) glasSummary[tier] = { count: 0, shared: 0, webOnly: 0, vdmSum: 0, netSum: 0 };
        
        glasSummary[tier].count++;
        if (r[dIdx["Fulfillment Tag"]] === "SHARED") glasSummary[tier].shared++;
        if (r[dIdx["Fulfillment Tag"]] === "WEBONLY") glasSummary[tier].webOnly++;
        glasSummary[tier].vdmSum += parseFloat(r[dIdx["VDM Markdown Depth %"]]) || 0;
        glasSummary[tier].netSum += (1 - (parseFloat(r[dIdx["Live Storefront Price"]]) / parseFloat(r[dIdx["Live Compare MSRP"]]))) || 0;
      }
    });

    const startRow = migrationRows.length + 4;
    const glasHeaders = [["House Strategic Tier", "Total SKUs", "Shared", "Web-Only", "% of Total", "VDM Discount %", "Affiliate Rate", "Final Stacked %"]];
    summarySheet.getRange(startRow, 1, 1, 8).setValues(glasHeaders).setBackground("#000000").setFontColor("#FFFFFF").setFontWeight("bold");

    const totalHouse = Object.values(glasSummary).reduce((a, b) => a + b.count, 0);
    const affiliateRate = ss.getSheetByName(VDM_CONFIG.TABS.CONTROL).getRange("E2").getValue();

    const glasRows = Object.keys(glasSummary).map(tier => {
      const t = glasSummary[tier];
      return [
        tier, t.count, t.shared, t.webOnly, totalHouse > 0 ? t.count / totalHouse : 0, 
        t.count > 0 ? t.vdmSum / t.count : 0, affiliateRate, t.count > 0 ? t.netSum / t.count : 0
      ];
    });

    if (glasRows.length > 0) {
      summarySheet.getRange(startRow + 1, 1, glasRows.length, 8).setValues(glasRows);
      summarySheet.getRange(startRow + 1, 5, glasRows.length, 4).setNumberFormat("0.00%");
    }

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

function buildStorefrontSyncAudit(dashRows, dIdx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const syncSheet = getOrCreateSheet(VDM_CONFIG.TABS.SYNC);
  syncSheet.clear(); // Clear before writing

  const shopifyRaw = ss.getSheetByName(VDM_CONFIG.TABS.RAW_SHOPIFY);
  const shopifyData = shopifyRaw.getDataRange().getValues();
  const sIdx = getHeaderMap(shopifyData[0]);

  const handleMap = {};
  shopifyData.slice(1).forEach(r => handleMap[sanitizeKey(r[sIdx["Variant SKU"]])] = r[sIdx["Handle"]]);

  // Ensure dashRows is not empty before mapping
  if (dashRows.length === 0) return;

  const syncAuditData = dashRows.map(r => {
    const sku = r[dIdx["SKU Anchor Key"]];
    const discountDepth = r[dIdx["VDM Markdown Depth %"]];
    const status = r[dIdx["Pricing Migration Status"]] || "";
    return [
      sku,                                // Col 1: SKU
      handleMap[sku] || "",               // Col 2: Handle
      status.includes("✓") ? "NO CHANGE" : "UPDATED", // Col 3: Action
      r[dIdx["Target Strategic Tier"]],   // Col 4: Final Tier
      discountDepth,                      // Col 5: Final Discount
      r[dIdx["Live Storefront Price"]],   // Col 6: Old Variant Price
      r[dIdx["Live Compare MSRP"]],       // Col 7: Old Compare At Price
      r[dIdx["Live Compare MSRP"]],       // Col 8: Base Price Used
      r[dIdx["New Proposed Storefront Price"]], // Col 9: New Variant Price
      discountDepth > 0 ? r[dIdx["Live Compare MSRP"]] : "", // Col 10: New Compare At Price
      r[dIdx["Profit Guardrail Status Alert"]] // Col 11: Note / Guardrail
    ];
  });

  const headers = [["SKU", "Handle", "Action", "Final Tier", "Final Discount", "Old Price", "Old Compare", "Base MSRP", "New Price", "New Compare", "Note"]];
  syncSheet.getRange(1, 1, 1, 11).setValues(headers);
  applyHeaderStyle(syncSheet.getRange(1, 1, 1, 11)); // Fix 3: Use applyHeaderStyle
  syncSheet.getRange(2, 1, syncAuditData.length, 11).setValues(syncAuditData);
  
  // Note/Guardrail Highlighting
  const range = syncSheet.getRange(2, 11, syncAuditData.length, 1);
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("❌ BLOCKED")
    .setBackground(VDM_CONFIG.DESIGN.ALERT_BREACH_BG)
    .setFontColor(VDM_CONFIG.DESIGN.ALERT_BREACH_TEXT)
    .setRanges([range])
    .build();
  syncSheet.setConditionalFormatRules([rule]);
}

/**
 * NEW: Generates the Master Pricing & Margin Ledger presentation tab.
 */
function generateMasterLedgerTab(dashRows, dIdx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ledgerSheet = getOrCreateSheet(VDM_CONFIG.TABS.MASTER_LEDGER);
  ledgerSheet.clear();
  ledgerSheet.clearConditionalFormatRules();

  const shopifyRaw = ss.getSheetByName(VDM_CONFIG.TABS.RAW_SHOPIFY);
  const shopifyData = shopifyRaw.getDataRange().getValues();
  const sIdx = getHeaderMap(shopifyData[0]);

  const handleMap = {};
  shopifyData.slice(1).forEach(r => handleMap[sanitizeKey(r[sIdx["Variant SKU"]])] = r[sIdx["Handle"]]);

  if (dashRows.length === 0) return;

  const ledgerData = dashRows.map(r => {
    const sku = r[dIdx["SKU Anchor Key"]];
    const status = r[dIdx["Pricing Migration Status"]] || "";
    const discountDepth = r[dIdx["VDM Markdown Depth %"]];
    return [
      sku,                                            // Col 1: SKU
      handleMap[sku] || "",                           // Col 2: Handle
      r[dIdx["Gatekeeper Status"]],                   // Col 3: Gatekeeper Status
      r[dIdx["Target Strategic Tier"]],               // Col 4: Final Tier
      status.includes("✓") ? "NO CHANGE" : "UPDATED", // Col 5: Action
      r[dIdx["Live Storefront Price"]],               // Col 6: Old Variant Price
      r[dIdx["Live Compare MSRP"]],                   // Col 7: Old Compare At Price
      r[dIdx["Live Compare MSRP"]],                   // Col 8: Base Price Used
      discountDepth,                                  // Col 9: Final Discount %
      r[dIdx["New Proposed Storefront Price"]],       // Col 10: New Variant Price
      discountDepth > 0 ? r[dIdx["Live Compare MSRP"]] : "", // Col 11: New Compare At Price
      r[dIdx["Simulated Checkout Net Price"]],        // Col 12: Simulated Checkout Net Price
      r[dIdx["Resolved Cost Base"]],                  // Col 13: Resolved Cost Base
      r[dIdx["Final Simulated Stacked Margin %"]],    // Col 14: Final Stacked Margin %
      r[dIdx["Profit Guardrail Status Alert"]]        // Col 15: Guardrail Alert
    ];
  });

  const headers = [["SKU", "Handle", "Gatekeeper Status", "Final Tier", "Action", "Old Variant Price", "Old Compare At Price", "Base Price Used", "Final Discount %", "New Variant Price", "New Compare At Price", "Simulated Checkout Net Price", "Resolved Cost Base", "Final Stacked Margin %", "Guardrail Alert"]];
  ledgerSheet.getRange(1, 1, 1, 15).setValues(headers);
  applyHeaderStyle(ledgerSheet.getRange(1, 1, 1, 15));
  ledgerSheet.getRange(2, 1, ledgerData.length, 15).setValues(ledgerData);

  const range = ledgerSheet.getRange(2, 1, ledgerData.length, 15);
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$O2="❌ BLOCKED"')
    .setBackground(VDM_CONFIG.DESIGN.ALERT_BREACH_BG)
    .setFontColor(VDM_CONFIG.DESIGN.ALERT_BREACH_TEXT)
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
      const itemSku = sanitizeKey(r[sIdx["Variant SKU"]]);
      if (itemSku) vendorMap[itemSku] = r[sIdx["Vendor"]];
    });

    // --- Execute all modules with correct arguments ---
    generateExecutiveSummary(rows, dIdx);
    buildStorefrontSyncAudit(rows, dIdx);
    logPricingElasticitySnapshot(rows, dIdx);
    generateSupplierScorecard(rows, vendorMap);
    generateWarehouseAgingReport(rows);
    generateMAPComplianceReport(rows);
    generateMasterLedgerTab(rows, dIdx);

    ui.alert("VDM v2.2 Refresh Cycle Complete. Strategic Tiers Updated.");
  } catch (e) {
    logError("MainCycle", e);
    ui.alert("ERROR: " + e.message);
  }
}