/**
 * MODULE 4: ANALYTICS REPORTING
 */

function generateAllReports(dashboardState) {
  try {
    if (!dashboardState || !dashboardState.rows) throw new Error("Dashboard state missing for reporting.");
    const { rows, headers } = dashboardState;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rawShopifySheet = ss.getSheetByName(VDM_CONFIG.TABS.RAW_SHOPIFY);
    if (!rawShopifySheet) throw new Error("Shopify raw data missing.");

    // Load Shopify Memory Map for Handle and Vendor lookups
    const shopifyRaw = rawShopifySheet.getDataRange().getValues();
    const sRawIdx = getHeaderMap(shopifyRaw[0]); // Dynamic index mapping
    const shopifyMap = new Map(shopifyRaw.slice(1).map(r => {
      // Map SKU Anchor to Handle and Vendor metadata safely
      return [r[0], { handle: r[sRawIdx["HANDLE"]], vendor: r[sRawIdx["VENDOR"]] }];
    }));

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

function generateSummaryTab(ss, rows, idx, shopifyMap) {
  const sheet = ss.getSheetByName(VDM_CONFIG.TABS.SUMMARY);
  sheet.clear().clearFormats();

  // Block 1: Migration Matrix
  const matrix = {};
  const tiers = ["Top Hero", "Signature Hero", "Proven Performer", "Accelerator", "Clearance", "New Launch", "Archive"];
  
  rows.forEach(r => {
    const from = r[idx["CURRENT EQUIVALENT STOREFRONT TIER"]];
    const to = r[idx["TARGET STRATEGIC TIER"]];
    if (!from || !to) return;
    
    const cost = safeNum(r[idx["RESOLVED COST BASE"]]) || 0;
    // Extract clean name (remove % Off text) for the key, ensuring it's a string
    const cleanTo = to.split(" (")[0];
    const key = `${from} -> ${cleanTo}`;

    if (!matrix[key]) matrix[key] = { count: 0, val: 0 };
    matrix[key].count++;
    matrix[key].val += cost;
  });

  const out = [["Migration Path", "SKU Count", "Invoiced Cost Value"]];
  Object.keys(matrix).forEach(k => out.push([k, matrix[k].count, matrix[k].val]));
  if (out.length > 0) {
    sheet.getRange(1, 1, out.length, 3).setValues(out);
    applyHeaderStyle(sheet.getRange(1, 1, 1, 3));
  }

  // Block 2: House Brands
  const houseRows = rows.filter(r => {
    const sku = r[idx["SKU ANCHOR KEY"]];
    const vendorName = (shopifyMap.get(sku)?.vendor || "").toUpperCase();
    return VDM_CONFIG.HOUSE_BRANDS.some(hb => vendorName.includes(hb.toUpperCase()));
  });

  // Fetch the exact global affiliate rate from your Settings tab so it applies universally
  const settingsData = ss.getSheetByName(VDM_CONFIG.TABS.SETTINGS).getDataRange().getValues();
  const globalAffiliateRate = (settingsData.length > 1 && safeNum(settingsData[1][4]) !== null) ? safeNum(settingsData[1][4]) : 0.15;

  // Block 2: House Brand Strategic Distribution
  const houseTiers = {};
  const totalHouseSKUs = houseRows.length;

  houseRows.forEach(hr => {
    const tierRaw = hr[idx["TARGET STRATEGIC TIER"]];
    if (!tierRaw) return;
    const tier = tierRaw.split(" (")[0];
    const mkdn = safeNum(hr[idx["VDM MARKDOWN DEPTH %"]]) || 0;
    const fulfillment = hr[idx["FULFILLMENT TAG"]];

    if (!houseTiers[tier]) {
      houseTiers[tier] = { count: 0, shared: 0, webOnly: 0, baseDiscount: mkdn };
    }
    
    houseTiers[tier].count++;
    if (fulfillment === "SHARED") houseTiers[tier].shared++;
    if (fulfillment === "WEBONLY") houseTiers[tier].webOnly++;
  });

  const houseOut = [["House Strategic Tier", "Total SKUs Count", "Shared Allocation", "Web-Only Allocation", "% of Total House Items", "VDM Base Discount %", "Affiliate Coupon Rate", "Final Stacked Checkout Discount"]];
  
  Object.keys(houseTiers).forEach(t => {
    const d = houseTiers[t];
    const pctOfTotal = totalHouseSKUs > 0 ? (d.count / totalHouseSKUs) : 0;
    // Mathematically compound the base discount with the global affiliate rate
    const finalStacked = 1 - ((1 - d.baseDiscount) * (1 - globalAffiliateRate));
    
    houseOut.push([
      t, 
      d.count, 
      d.shared, 
      d.webOnly, 
      pctOfTotal, 
      d.baseDiscount, 
      globalAffiliateRate, 
      finalStacked
    ]);
  });

  const startRow = out.length + 4;
  sheet.getRange(startRow - 1, 1).setValue("House Brand Analytics").setFontWeight("bold");
  if (houseOut.length > 1) {
    sheet.getRange(startRow, 1, houseOut.length, 8).setValues(houseOut);
    applyHeaderStyle(sheet.getRange(startRow, 1, 1, 8));
    sheet.getRange(startRow + 1, 5, houseOut.length - 1, 1).setNumberFormat("0.00%");
    sheet.getRange(startRow + 1, 6, houseOut.length - 1, 3).setNumberFormat("0.00%");
  }
}

function generateSyncAudit(ss, rows, idx, shopifyMap) {
  const sheet = ss.getSheetByName(VDM_CONFIG.TABS.SYNC_AUDIT);
  sheet.clear().clearFormats();

  const headers = ["SKU", "Handle", "Action", "Final Tier", "Final Discount", "Old Variant Price", "Old Compare At Price", "Base Price Used", "New Variant Price", "New Compare At Price", "Note"];

  const syncRows = rows.map(r => {
    const sku = r[idx["SKU ANCHOR KEY"]];
    const mkdn = safeNum(r[idx["VDM MARKDOWN DEPTH %"]]) || 0;
    
    // Terminology Alignment (Manual Section 4.4): Standardize Action keywords
    const status = r[idx["PRICING MIGRATION STATUS"]];
    const action = (status === "✓ Price Hold" || status === "⚠️ HOLD: B2B Volume Stable") ? "NO CHANGE" : "UPDATED";

    return [
      sku,
      shopifyMap.get(sku)?.handle || "",
      action,
      r[idx["TARGET STRATEGIC TIER"]],
      mkdn,
      r[idx["LIVE STOREFRONT PRICE"]],
      r[idx["LIVE COMPARE MSRP"]],
      r[idx["LIVE COMPARE MSRP"]],
      r[idx["NEW PROPOSED STOREFRONT PRICE"]],
      mkdn === 0 ? "" : r[idx["LIVE COMPARE MSRP"]],
      ""
    ];
  });

  if (headers.length > 0) {
    sheet.getRange(1, 1, 1, 11).setValues([headers]);
    applyHeaderStyle(sheet.getRange(1, 1, 1, 11));
  }
  if (syncRows.length > 0) sheet.getRange(2, 1, syncRows.length, 11).setValues(syncRows);
}

function generateMasterLedger(ss, rows, idx, shopifyMap) {
  const sheet = ss.getSheetByName(VDM_CONFIG.TABS.MASTER_LEDGER);
  sheet.clear().clearFormats();

  const headers = ["SKU", "Handle", "Gatekeeper Status", "Final Tier", "Action", "Old Variant Price", "Old Compare At Price", "Base Price Used", "Final Discount %", "New Variant Price", "New Compare At Price", "Simulated Checkout Net Price", "Resolved Cost Base", "Final Stacked Margin %", "Guardrail Alert"];
  
  const ledgerRows = rows.map(r => {
    const sku = r[idx["SKU ANCHOR KEY"]];
    const mkdn = safeNum(r[idx["VDM MARKDOWN DEPTH %"]]) || 0;
    return [
      sku,
      shopifyMap.get(sku)?.handle || "",
      r[idx["GATEKEEPER STATUS"]],
      r[idx["TARGET STRATEGIC TIER"]],
      r[idx["PRICING MIGRATION STATUS"]],
      r[idx["LIVE STOREFRONT PRICE"]],
      r[idx["LIVE COMPARE MSRP"]],
      r[idx["LIVE COMPARE MSRP"]],
      mkdn,
      r[idx["NEW PROPOSED STOREFRONT PRICE"]],
      mkdn === 0 ? "" : r[idx["LIVE COMPARE MSRP"]],
      r[idx["SIMULATED CHECKOUT NET PRICE"]],
      r[idx["RESOLVED COST BASE"]],
      r[idx["FINAL SIMULATED STACKED MARGIN %"]],
      r[idx["PROFIT GUARDRAIL STATUS ALERT"]]
    ];
  });

  if (headers.length > 0) {
    sheet.getRange(1, 1, 1, 15).setValues([headers]);
    applyHeaderStyle(sheet.getRange(1, 1, 1, 15));
  }
  if (ledgerRows.length > 0) sheet.getRange(2, 1, ledgerRows.length, 15).setValues(ledgerRows);
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

function executeFlexibleRefreshProcess() {
  const ui = SpreadsheetApp.getUi();
  try {
    ui.showModelessDialog(HtmlService.createHtmlOutput("<b>Processing VDM Refresh...</b>"), "System Status");
    runDataIngestion();
    const dashboardState = executeDashboardRefresh();
    generateAllReports(dashboardState);
    ui.alert("VDM Refresh Complete.");
  } catch (e) {
    ui.alert("Process Failed: " + e.message);
  }
}