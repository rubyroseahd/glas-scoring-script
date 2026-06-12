/**
 * MODULE 4: ANALYTICS REPORTING
 * Executive Summaries, Presentation Dashboards & Sync Audits
 */

function generateExecutiveSummary() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dash = ss.getSheetByName(VDM_CONFIG.TABS.DASHBOARD);
    const dashData = dash.getDataRange().getValues();
    const dashHeaders = dashData[0];
    const dIdx = getHeaderMap(dashHeaders);
    const rows = dashData.slice(1);

    // Block 1: Global Tier Migration Matrix
    const migrationMatrix = {};
    rows.forEach(r => {
      // Dynamically find values using headers
      const from = r[dIdx["Current Tier"]] || "N/A";
      const to = r[dIdx["Target Tier"]]; 
      const cost = parseFloat(r[dIdx["Invoiced Asset Cost"]]) || 0; 
      
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
    const vendorIdx = shopHeaders.findIndex(h => h.includes("Vendor"));
    const skuIdx = shopHeaders.findIndex(h => h.includes("Variant SKU"));

    const vendorMap = {};
    shopifyData.slice(1).forEach(r => vendorMap[sanitizeKey(r[skuIdx])] = r[vendorIdx]);

    const glasSummary = {};
    rows.forEach(r => {
      const sku = r[dIdx["SKU"]];
      const vendor = (vendorMap[sku] || "").toLowerCase();
      const isHouseBrand = VDM_CONFIG.HOUSE_BRANDS.some(hb => vendor.includes(hb.toLowerCase()));

      if (isHouseBrand) {
        const tier = r[dIdx["Target Tier"]];
        if (!glasSummary[tier]) glasSummary[tier] = { count: 0, shared: 0, webOnly: 0, vdmSum: 0, netSum: 0 };
        
        glasSummary[tier].count++;
        if (r[dIdx["Status"]] === "SHARED") glasSummary[tier].shared++;
        if (r[dIdx["Status"]] === "WEBONLY") glasSummary[tier].webOnly++;
        glasSummary[tier].vdmSum += parseFloat(r[dIdx["Discount Depth"]]) || 0;
        glasSummary[tier].netSum += (1 - (parseFloat(r[dIdx["Price"]]) / parseFloat(r[dIdx["MSRP"]]))) || 0;
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

    buildStorefrontSyncAudit(rows, dIdx);
    logPricingElasticitySnapshot(rows, dIdx);
    generateSupplierScorecard(rows, vendorMap);
    generateWarehouseAgingReport(rows);
    generateMAPComplianceReport(rows);

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
    r[dIdx["SKU"]], 
    r[dIdx["Discount Depth"]], 
    r[dIdx["Price"]], 
    r[dIdx["Velocity Score"]]
  ]);
  if (logData.length > 0) {
    const lastRow = ledger.getLastRow();
    ledger.getRange(lastRow === 0 ? 1 : lastRow + 1, 1, logData.length, 5).setValues(logData);
    if (lastRow === 0) {
      ledger.insertRowBefore(1);
      ledger.getRange("A1:E1").setValues([["Date", "SKU", "Markdown %", "Checkout Price", "Velocity Score"]]).setFontWeight("bold");
    }
  }
}

function generateSupplierScorecard(rows, vendorMap) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(VDM_CONFIG.TABS.SCORECARD);
  sheet.getRange("A1").setValue("Supplier Scorecard Summary (WIP)").setFontWeight("bold");
}

function generateWarehouseAgingReport(rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(VDM_CONFIG.TABS.AGING);
  sheet.getRange("A1").setValue("Liquidation Candidates (WIP)").setFontWeight("bold");
}

function generateMAPComplianceReport(rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(VDM_CONFIG.TABS.COMPLIANCE);
  sheet.getRange("A1").setValue("MAP Compliance Tracking (WIP)").setFontWeight("bold");
}

function buildStorefrontSyncAudit(dashRows, dIdx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shopifyRaw = ss.getSheetByName(VDM_CONFIG.TABS.RAW_SHOPIFY);
  const shopifyData = shopifyRaw.getDataRange().getValues();
  const shopHeaders = shopifyData[0];
  const handleIdx = shopHeaders.findIndex(h => h === "Handle");
  const skuIdx = shopHeaders.findIndex(h => h.includes("Variant SKU"));

  const handleMap = {};
  shopifyData.slice(1).forEach(r => handleMap[sanitizeKey(r[skuIdx])] = r[handleIdx]);

  const syncAuditData = dashRows.map(r => {
    const sku = r[dIdx["SKU"]];
    const discountDepth = r[dIdx["Discount Depth"]];
    const status = r[dIdx["Update Status"]] || "";
    
    return [
      sku,                                // Col 1: SKU
      handleMap[sku] || "",               // Col 2: Handle
      status.includes("✓") ? "NO CHANGE" : "UPDATED", // Col 3: Action
      r[dIdx["Target Tier"]],             // Col 4: Final Tier
      discountDepth,                      // Col 5: Final Discount
      r[dIdx["Old Price"]],               // Col 6: Old Variant Price
      r[dIdx["Old MSRP"]],                // Col 7: Old Compare At Price
      r[dIdx["Old MSRP"]],                // Col 8: Base Price Used
      r[dIdx["Price"]],                   // Col 9: New Variant Price
      discountDepth > 0 ? r[dIdx["Old MSRP"]] : "", // Col 10: New Compare At Price
      r[dIdx["Guardrail Note"]]           // Col 11: Note / Guardrail
    ];
  });

  const syncSheet = getOrCreateSheet(VDM_CONFIG.TABS.SYNC);
  const headers = [["SKU", "Handle", "Action", "Final Tier", "Final Discount", "Old Price", "Old Compare", "Base MSRP", "New Price", "New Compare", "Note"]];
  syncSheet.getRange(1, 1, 1, 11).setValues(headers).setBackground("#000000").setFontColor("#FFFFFF").setFontWeight("bold");
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

function runFullRefreshCycle() {
  const ui = SpreadsheetApp.getUi();
  try {
    runDataIngestion();
    refreshDashboardMatrix();
    generateExecutiveSummary();
    ui.alert("VDM v2.2 Refresh Cycle Complete. Strategic Tiers Updated.");
  } catch (e) {
    logError("MainCycle", e);
    ui.alert("ERROR: " + e.message);
  }
}