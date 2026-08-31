/**
 * MODULE 4: ANALYTICS REPORTING
 */

function generateAllReports(dashboardState) {
  const reportState = {
    success: false,
    failedReports: [],
    errorDetails: "",
    summaryTelemetry: null
  };
  const failureDetails = [];
  const recordFailure = (reportName, err) => {
    const message = safeStr(err && err.message) || safeStr(err) || "Unknown error";
    reportState.failedReports.push(reportName);
    failureDetails.push(`${reportName}: ${message}`);
    logError("Reporting", err instanceof Error ? err : new Error(`${reportName} failed: ${message}`));
  };

  try {
    if (!dashboardState || !Array.isArray(dashboardState.rows) || !Array.isArray(dashboardState.headers)) {
      recordFailure("Dashboard State", new Error("Dashboard state missing or invalid for reporting."));
      reportState.errorDetails = failureDetails.join(" | ");
      return reportState;
    }
    const { rows, headers } = dashboardState;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shopifyMap = getShopifyMap();
    if (shopifyMap.size === 0) {
      recordFailure("Shopify Metadata", new Error("Shopify metadata missing. Please run ingestion."));
      reportState.errorDetails = failureDetails.join(" | ");
      return reportState;
    }

    const idx = getHeaderMap(headers); // Use the standardized helper for Dashboard columns

    const reportRunners = [
      {
        name: "Executive Summary",
        run: () => {
          reportState.summaryTelemetry = generateSummaryTab(ss, rows, idx, shopifyMap);
        }
      },
      { name: "Action Items", run: () => generateActionItems(ss, rows, idx, shopifyMap) },
      { name: "Sync Audit", run: () => generateSyncAudit(ss, rows, idx, shopifyMap) },
      { name: "Master Ledger", run: () => generateMasterLedger(ss, rows, idx, shopifyMap) },
      { name: "Supplier Scorecard", run: () => generateSupplierScorecard(ss, rows, idx, shopifyMap) },
      { name: "Elasticity Snapshot", run: () => logElasticitySnapshot(ss, rows, idx) }
    ];

    reportRunners.forEach(report => {
      try {
        report.run();
      } catch (e) {
        recordFailure(report.name, e);
      }
    });

    reportState.success = reportState.failedReports.length === 0;
    reportState.errorDetails = reportState.success ? "" : failureDetails.join(" | ");
    return reportState;
  } catch (e) {
    logError("Reporting", e);
    recordFailure("Reporting Pipeline", e);
    reportState.errorDetails = failureDetails.join(" | ");
    return reportState;
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
  const handleHeader = findFirstAvailableHeader(idx, ["HANDLE"]);
  const vendorHeader = findFirstAvailableHeader(idx, ["VENDOR"]);
  return new Map(data.slice(1).map(r => [
    r[0],
    {
      handle: handleHeader ? r[idx[handleHeader]] : "",
      vendor: vendorHeader ? r[idx[vendorHeader]] : ""
    }
  ]));
}

function getDashboardReportHeaders(idx) {
  return {
    livePrice: findFirstAvailableHeader(idx, ["VARIANT PRICE", "LIVE STOREFRONT PRICE"]),
    compareAtPrice: findFirstAvailableHeader(idx, ["COMPARE AT PRICE", "LIVE COMPARE MSRP"]),
    proposedPrice: findFirstAvailableHeader(idx, ["PROPOSED VARIANT PRICE", "NEW PROPOSED STOREFRONT PRICE"]),
    unitsSold90: findFirstAvailableHeader(idx, ["90-DAY NET ITEMS SOLD", "RAW 90D RETAIL VELOCITY"]),
    currentDiscount: findFirstAvailableHeader(idx, ["CURRENT DISCOUNT %", "ACTIVE STOREFRONT MARKDOWN DEPTH %"]),
    currentMargin: findFirstAvailableHeader(idx, ["CURRENT MARGIN %", "CURRENT GROSS MARGIN %"]),
    projectedMargin: findFirstAvailableHeader(idx, ["PROJECTED MARGIN %", "FINAL SIMULATED STACKED MARGIN %"]),
    unitCost: findFirstAvailableHeader(idx, ["UNIT COST", "RESOLVED COST BASE"]),
    simulatedNetPrice: findFirstAvailableHeader(idx, ["SIMULATED NET PRICE", "SIMULATED CHECKOUT NET PRICE"]),
    marginDelta: findFirstAvailableHeader(idx, ["MARGIN DELTA %", "NET MARGIN CHANGE %"]),
    totalNetworkStock: findFirstAvailableHeader(idx, ["TOTAL NETWORK STOCK", "TOTAL ON-HAND WAREHOUSE STOCK"]),
    priceShift: findFirstAvailableHeader(idx, ["PRICE SHIFT ($)", "RETAIL PRICE SHIFT ($)"])
  };
}

function generateSummaryTab(ss, rows, idx, shopifyMap) {
  const sheet = ss.getSheetByName(VDM_CONFIG.TABS.SUMMARY) || ss.insertSheet(VDM_CONFIG.TABS.SUMMARY);
  sheet.clear().clearFormats();
  const dashboardHeaders = getDashboardReportHeaders(idx);
  
  const settingsData = ss.getSheetByName(VDM_CONFIG.TABS.SETTINGS).getDataRange().getValues();
  const globalAffiliateRate = (settingsData.length > 1 && safeNum(settingsData[1][4]) !== null) ? safeNum(settingsData[1][4]) : 0.15;

  // --- PANEL A: GLOBAL CATALOG COMPARATIVE DISTRIBUTION MATRIX ---
  const panelAHeaders = ["Storefront Bracket", "Baseline SKU Count", "Baseline Catalog %", "VDM Proposed SKU Count", "VDM Proposed Catalog %", "Net SKU Shift", "Net Catalog Shift %"];
  const bracketOrder = [
    VDM_CONFIG.BRACKET_NAMES.HERO,
    VDM_CONFIG.BRACKET_NAMES.SIGNATURE,
    VDM_CONFIG.BRACKET_NAMES.PROVEN,
    VDM_CONFIG.BRACKET_NAMES.ACCELERATOR,
    VDM_CONFIG.BRACKET_NAMES.CLEARANCE
  ];

  const baselineBracketHeader = findFirstAvailableHeader(idx, ["BASELINE STOREFRONT BRACKET"]);
  const targetBracketHeader = findFirstAvailableHeader(idx, ["TARGET VDM BRACKET"]);
  const livePriceHeader = dashboardHeaders.livePrice;
  const baselinePriceHeader = findFirstAvailableHeader(idx, ["BASELINE STOREFRONT PRICE"]);
  const baselineSkuPresentHeader = findFirstAvailableHeader(idx, ["BASELINE SKU PRESENT"]);
  const compareHeader = dashboardHeaders.compareAtPrice;
  const proposedPriceHeader = dashboardHeaders.proposedPrice;
  const vdmMarkdownHeader = findFirstAvailableHeader(idx, ["VDM MARKDOWN DEPTH %"]);
  const unitsHeader = dashboardHeaders.unitsSold90;

  const totalRows = rows.length;
  const baselineCounts = {};
  const targetCounts = {};
  bracketOrder.forEach(name => {
    baselineCounts[name] = 0;
    targetCounts[name] = 0;
  });

  let clearanceCount = 0;
  let projectedRevenueImpact90 = 0;

  rows.forEach(r => {
    const livePrice = livePriceHeader ? safeNum(r[idx[livePriceHeader]]) : null;
    const baselinePrice = baselinePriceHeader ? safeNum(r[idx[baselinePriceHeader]]) : livePrice;
    const baselineSkuPresent = baselineSkuPresentHeader ? isBaselineSkuPresentValue(r[idx[baselineSkuPresentHeader]]) : false;
    const compareMsrp = compareHeader ? safeNum(r[idx[compareHeader]]) : null;
    const proposedPrice = proposedPriceHeader ? safeNum(r[idx[proposedPriceHeader]]) : null;
    const units90 = unitsHeader ? (safeNum(r[idx[unitsHeader]]) || 0) : 0;
    const rowMarkdown = vdmMarkdownHeader ? safeNum(r[idx[vdmMarkdownHeader]]) : null;

    const fallbackBaselineBracket = getStorefrontBracket(
      baselinePrice !== null ? baselinePrice : (livePrice || 0),
      compareMsrp !== null ? compareMsrp : (baselinePrice !== null ? baselinePrice : (livePrice || 0))
    );
    const fallbackTargetBracket = getStorefrontBracket(
      proposedPrice !== null ? proposedPrice : (livePrice || 0),
      compareMsrp !== null ? compareMsrp : (proposedPrice !== null ? proposedPrice : (livePrice || 0))
    );

    const _canonicalBrackets = Object.values(VDM_CONFIG.BRACKET_NAMES);
    const _rawBaselineBracket = baselineBracketHeader ? (safeStr(r[idx[baselineBracketHeader]]) || null) : null;
    const _rawTargetBracket = targetBracketHeader ? (safeStr(r[idx[targetBracketHeader]]) || null) : null;
    const baselineBracket = (_rawBaselineBracket && _canonicalBrackets.indexOf(_rawBaselineBracket) !== -1)
      ? _rawBaselineBracket
      : fallbackBaselineBracket;
    const targetBracket = (_rawTargetBracket && _canonicalBrackets.indexOf(_rawTargetBracket) !== -1)
      ? _rawTargetBracket
      : fallbackTargetBracket;

    if (baselineCounts[baselineBracket] !== undefined) baselineCounts[baselineBracket]++;
    if (targetCounts[targetBracket] !== undefined) targetCounts[targetBracket]++;
    if ((rowMarkdown || 0) >= 0.65) clearanceCount++;

    projectedRevenueImpact90 += getProjectedRevenueImpact90Contribution(
      baselineSkuPresent,
      proposedPrice,
      baselinePrice,
      units90
    );
  });

  const panelAData = bracketOrder.map(bracketName => {
    const baselineCount = baselineCounts[bracketName] || 0;
    const targetCount = targetCounts[bracketName] || 0;
    const baselinePct = totalRows > 0 ? baselineCount / totalRows : 0;
    const targetPct = totalRows > 0 ? targetCount / totalRows : 0;
    const netSkuShift = targetCount - baselineCount;
    const netCatalogShiftPct = targetPct - baselinePct;
    return [bracketName, baselineCount, baselinePct, targetCount, targetPct, netSkuShift, netCatalogShiftPct];
  });

  const panelAWidth = panelAHeaders.length;
  sheet.getRange(1, 1).setValue("GLOBAL CATALOG ALLOCATION SUMMARY MATRIX").setFontSize(14).setFontWeight("bold").setBackground(VDM_CONFIG.DESIGN.PANEL_GLOBAL_BG).setFontColor("#FFFFFF");
  sheet.getRange(2, 1, 1, panelAWidth).setValues([panelAHeaders]);
  applyHeaderStyle(sheet.getRange(2, 1, 1, panelAWidth));
  sheet.getRange(3, 1, panelAData.length, panelAWidth).setValues(panelAData);

  if (panelAData.length > 0) {
    const panelAHeaderMap = getHeaderMap(panelAHeaders);
    ["Baseline Catalog %", "VDM Proposed Catalog %", "Net Catalog Shift %"].forEach(header => {
      const columnIndex = panelAHeaderMap[safeStr(header).toUpperCase()];
      if (columnIndex !== undefined) {
        sheet.getRange(3, columnIndex + 1, panelAData.length, 1).setNumberFormat("0.00%");
      }
    });
    const netSkuShiftIndex = panelAHeaderMap[safeStr("Net SKU Shift").toUpperCase()];
    if (netSkuShiftIndex !== undefined) {
      sheet.getRange(3, netSkuShiftIndex + 1, panelAData.length, 1).setNumberFormat("0");
    }
  }

  // --- PANEL B: GLÄS & GLASTOY PROPRIETARY CATALOG DELTA PANEL ---
  const proprietaryBrackets = [
    { name: "Top Hero Bracket", mkdn: 0.00, shopCheck: (m) => m === 0, vdmMatch: "Top Hero" },
    { name: "Signature Hero Bracket", mkdn: 0.30, shopCheck: (m) => m > 0 && m <= 0.35, vdmMatch: "Signature Hero" },
    { name: "Proven Performer Bracket", mkdn: 0.40, shopCheck: (m) => m > 0.35 && m <= 0.45, vdmMatch: "Proven Performer" },
    { name: "Accelerator Bracket", mkdn: 0.50, shopCheck: (m) => m > 0.45 && m <= 0.55, vdmMatch: "Accelerator" },
    { name: "Clearance/Archive Bracket", mkdn: 0.65, shopCheck: (m) => m > 0.55, vdmMatch: "Clearance/Archive" },
    { name: "New Launch Bracket", mkdn: 0.00, shopCheck: (m) => false, vdmMatch: "New Launch" },
    { name: "B2B Protection Hold Bracket", mkdn: 0.00, shopCheck: (m) => false, vdmMatch: "B2B Protection Hold" }
  ];
  const houseRows = rows.filter(r => {
    const sku = r[idx["SKU ANCHOR KEY"]];
    const vendorName = (shopifyMap.get(sku)?.vendor || "").toUpperCase();
    return VDM_CONFIG.HOUSE_BRANDS.some(hb => vendorName.includes(hb.toUpperCase()));
  });
  const totalHouseRows = houseRows.length;
  const panelBHeaders = ["Proprietary Bracket", "Shopify Count", "Shopify %", "VDM Count", "VDM %", "Shift %", "Base %", "Stacked %"];
  
  let panelBData = totalHouseRows === 0 ? [] : proprietaryBrackets.map(b => {
    const shopCount = houseRows.filter(r => b.shopCheck(safeNum(r[idx[dashboardHeaders.currentDiscount]]) ?? 0)).length;
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
  sheet.getRange(startPanelB + 1 + panelBData.length, 1, 1, panelBWidth).setFontWeight("bold");

  [3, 5, 6, 7, 8].forEach(col => {
    if (panelBData.length > 0) sheet.getRange(startPanelB + 2, col, panelBData.length, 1).setNumberFormat("0.00%");
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

  const clearanceRatio = totalRows > 0 ? clearanceCount / totalRows : 0;
  const clearanceCapWarn = clearanceRatio > VDM_CONFIG.CLEARANCE_CAP_WARN;
  if (clearanceCapWarn) {
    Logger.log(`[EXEC WARNING] Clearance catalog ratio ${clearanceRatio.toFixed(4)} exceeds threshold ${VDM_CONFIG.CLEARANCE_CAP_WARN.toFixed(4)}.`);
  }

  const baselineHeroCount = baselineCounts[VDM_CONFIG.BRACKET_NAMES.HERO] || 0;
  const targetHeroCount = targetCounts[VDM_CONFIG.BRACKET_NAMES.HERO] || 0;
  const heroShrinkage = baselineHeroCount > 0 ? (baselineHeroCount - targetHeroCount) / baselineHeroCount : 0;
  const heroDepletionAlert = baselineHeroCount > 0 && heroShrinkage > VDM_CONFIG.HERO_POOL_MIN_WARN;
  if (heroDepletionAlert) {
    Logger.log(`[EXEC WARNING] Hero pool shrinkage ${heroShrinkage.toFixed(4)} exceeds threshold ${VDM_CONFIG.HERO_POOL_MIN_WARN.toFixed(4)}.`);
  }

  const usedWidth = Math.max(panelAWidth, panelBWidth, 2);
  if (usedWidth > 0) sheet.autoResizeColumns(1, usedWidth);
  sheet.getRange(2, 1, 1, panelAWidth).setFontWeight("bold").setHorizontalAlignment("center");

  return {
    clearanceCount,
    totalRows,
    clearanceRatio,
    clearanceCapWarn,
    baselineHeroCount,
    targetHeroCount,
    heroShrinkage,
    heroDepletionAlert,
    projectedRevenueImpact90
  };
}

function isBaselineSkuPresentValue(value) {
  if (value === true) return true;
  if (value === false || value === null || value === undefined) return false;
  const normalized = safeStr(value).toUpperCase();
  return normalized === "TRUE";
}

function getProjectedRevenueImpact90Contribution(baselineSkuPresent, proposedPrice, baselinePrice, units90) {
  if (!baselineSkuPresent || !mathGuard(proposedPrice, baselinePrice)) return 0;
  return (proposedPrice - baselinePrice) * (safeNum(units90) || 0);
}

/**
 * Generates the Action Items & Sign-off tab with Queue 1A/1B/2/3 routing.
 * Queue 1A: Negative base margin audit
 * Queue 1B: Simulated checkout margin guardrail (<20%)
 * Queue 2: WEBONLY digital review (total score 0–3)
 * Queue 3: SHARED clearance/liquidation (total score 0–3)
 */
function generateActionItems(ss, rows, idx, shopifyMap) {
  const sheet = ss.getSheetByName(VDM_CONFIG.TABS.ACTION) || ss.insertSheet(VDM_CONFIG.TABS.ACTION);
  sheet.clear().clearFormats();
  const dashboardHeaders = getDashboardReportHeaders(idx);

  const headers = ["Queue", "SKU", "Handle", "Vendor", "Fulfillment", "Total Score", "Current Margin %", "Stacked Margin %", "Guardrail", "Target Tier", "Proposed Price", "Action Required"];
  const width = headers.length;
  sheet.getRange(1, 1, 1, width).setValues([headers]);
  applyHeaderStyle(sheet.getRange(1, 1, 1, width));

  const actionRows = rows
    .map(r => {
      const queue = r[idx["ACTION QUEUE"]] || "";
      if (!queue) return null;
      const sku = r[idx["SKU ANCHOR KEY"]];
      const meta = shopifyMap.get(sku) || {};
      const curMargin = safeNum(r[idx[dashboardHeaders.currentMargin]]) || 0;
      const stackMargin = safeNum(r[idx[dashboardHeaders.projectedMargin]]) || 0;
      let actionRequired = "";
      if (queue.startsWith("Queue 1A") && r[idx["QUEUE CODE"]] === QUEUE_CODES.QUEUE_1A_COST) {
        actionRequired = "Flag for cost audit — missing cost data";
      } else if (queue.startsWith("Queue 1A")) {
        actionRequired = "Flag for price/cost audit — base margin is negative";
      } else if (queue.startsWith("Queue 1B")) {
        actionRequired = "Pricing blocked — stacked margin below 20% guardrail";
      } else if (queue.startsWith("Queue 2")) {
        actionRequired = "WEBONLY digital review — score 0–3; route to digital clearance or archive";
      } else if (queue.startsWith("Queue 3")) {
        actionRequired = "SHARED clearance/liquidation — score 0–3; route to physical clearance channel";
      }
      return [
        queue, sku, meta.handle || "", meta.vendor || "",
        r[idx["FULFILLMENT TAG"]], r[idx["TOTAL COMPOSITE SCORE"]],
        curMargin, stackMargin, r[idx["PROFIT GUARDRAIL STATUS ALERT"]],
        r[idx["TARGET STRATEGIC TIER"]], r[idx[dashboardHeaders.proposedPrice]],
        actionRequired
      ];
    })
    .filter(r => r !== null);

  if (actionRows.length > 0) {
    sheet.getRange(2, 1, actionRows.length, width).setValues(actionRows);
    const headerMap = getHeaderMap(headers);
    ["Current Margin %", "Stacked Margin %"].forEach(header => {
      const columnIndex = headerMap[safeStr(header).toUpperCase()];
      if (columnIndex !== undefined) {
        sheet.getRange(2, columnIndex + 1, actionRows.length, 1).setNumberFormat("0.00%");
      }
    });
    const proposedPriceColumnIndex = headerMap["PROPOSED PRICE"];
    if (proposedPriceColumnIndex !== undefined) {
      sheet.getRange(2, proposedPriceColumnIndex + 1, actionRows.length, 1).setNumberFormat("$#,##0.00");
    }
  }
  sheet.setFrozenRows(1);
  applyActionQueueConditionalFormatting(sheet, 1);
}

function applyActionQueueConditionalFormatting(sheet, queueColumnIndex) {
  const queueRange = sheet.getRange(2, queueColumnIndex, Math.max(sheet.getMaxRows() - 1, 1), 1);
  const existingRules = sheet.getConditionalFormatRules().filter(rule => {
    return !rule.getRanges().some(range =>
      range.getSheet().getSheetId() === sheet.getSheetId() &&
      range.getColumn() <= queueColumnIndex &&
      range.getColumn() + range.getNumColumns() - 1 >= queueColumnIndex
    );
  });

  const queueRules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextStartsWith("Queue 1")
      .setBackground("#FCE8E6")
      .setFontColor("#A51D24")
      .setBold(true)
      .setRanges([queueRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextStartsWith("Queue 2")
      .setBackground("#FEF7E0")
      .setFontColor("#B06000")
      .setRanges([queueRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextStartsWith("Queue 3")
      .setBackground("#F1F3F4")
      .setFontColor("#3C4043")
      .setRanges([queueRange])
      .build()
  ];

  sheet.setConditionalFormatRules(existingRules.concat(queueRules));
}

function generateSyncAudit(ss, rows, idx, shopifyMap) {
  const sheet = ss.getSheetByName(VDM_CONFIG.TABS.SYNC_AUDIT) || ss.insertSheet(VDM_CONFIG.TABS.SYNC_AUDIT);
  sheet.clear().clearFormats();
  const dashboardHeaders = getDashboardReportHeaders(idx);

  // Revised Schema Layer 3
  const headers = ["SKU Key", "Handle", "Action", "Strategy Tier", "VDM Markdown %", "Old Price", "New Price", "Old MSRP", "New MSRP", "Base Price", "Guardrail", "Gatekeeper Status"];

  const syncRows = rows.map(r => {
    const sku = r[idx["SKU ANCHOR KEY"]];
    const handle = shopifyMap.get(sku)?.handle || "";
    const mkdn = safeNum(r[idx["VDM MARKDOWN DEPTH %"]]) || 0;
    const status = r[idx["PRICING MIGRATION STATUS"]];
    const action = (status === "✓ Price Hold" || status === "⚠️ HOLD: B2B Volume Stable")
      ? "NO CHANGE"
      : "UPDATED";
    const currentMsrp = r[idx[dashboardHeaders.compareAtPrice]];
    const nextMsrp = mkdn === 0 ? "" : currentMsrp;

    return [
      sku,
      handle,
      action,
      r[idx["TARGET STRATEGIC TIER"]],
      mkdn,
      r[idx[dashboardHeaders.livePrice]],
      r[idx[dashboardHeaders.proposedPrice]],
      currentMsrp,
      nextMsrp,
      currentMsrp,
      r[idx["PROFIT GUARDRAIL STATUS ALERT"]],
      r[idx["GATEKEEPER STATUS"]]
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

  return { headers, rows: syncRows };
}

function generateMasterLedger(ss, rows, idx, shopifyMap) {
  const sheet = getOrCreateSheet(VDM_CONFIG.TABS.MASTER_LEDGER);
  sheet.clear().clearFormats();
  const dashboardHeaders = getDashboardReportHeaders(idx);

  const headers = ["SKU Key", "Handle", "Fulfillment", "Gatekeeper", "Migration Status", "Tier", "Target Mkdn %", "Old Price", "New Price", "Price Shift ($)", "Procurement Cost", "Checkout Price", "Stacked Margin %", "Guardrail", "Margin Shift %"];
  
  const ledgerRows = rows.map(r => {
    return [
      r[idx["SKU ANCHOR KEY"]], shopifyMap.get(r[idx["SKU ANCHOR KEY"]])?.handle || "",
      r[idx["FULFILLMENT TAG"]], r[idx["GATEKEEPER STATUS"]],
      r[idx["PRICING MIGRATION STATUS"]], r[idx["TARGET STRATEGIC TIER"]], r[idx["VDM MARKDOWN DEPTH %"]],
      r[idx[dashboardHeaders.livePrice]], r[idx[dashboardHeaders.proposedPrice]], r[idx[dashboardHeaders.priceShift]],
      r[idx[dashboardHeaders.unitCost]], r[idx[dashboardHeaders.simulatedNetPrice]], r[idx[dashboardHeaders.projectedMargin]],
      r[idx["PROFIT GUARDRAIL STATUS ALERT"]], r[idx[dashboardHeaders.marginDelta]]
    ];
  });

  const width = headers.length;
  sheet.getRange(1, 1, 1, width).setValues([headers]);
  applyHeaderStyle(sheet.getRange(1, 1, 1, width));
  if (ledgerRows.length > 0) {
    const range = sheet.getRange(2, 1, ledgerRows.length, width);
    range.setValues(ledgerRows);

    // Phase 3 & 4 Financial Formatting
    [8, 9, 10, 11, 12].forEach(col => sheet.getRange(2, col, ledgerRows.length, 1).setNumberFormat("0.00"));
    [7, 13, 15].forEach(col => sheet.getRange(2, col, ledgerRows.length, 1).setNumberFormat("0.00%"));
  }
}

function generateSupplierScorecard(ss, rows, idx, shopifyMap) {
  const sheet = ss.getSheetByName(VDM_CONFIG.TABS.SCORECARD) || ss.insertSheet(VDM_CONFIG.TABS.SCORECARD);
  sheet.clear().clearFormats();
  const dashboardHeaders = getDashboardReportHeaders(idx);

  const vendorTotals = {};
  rows.forEach(r => {
    const sku = r[idx["SKU ANCHOR KEY"]];
    const vendor = shopifyMap.get(sku)?.vendor || "Unknown Vendor";
    const stockVal = (safeNum(r[idx[dashboardHeaders.totalNetworkStock]]) || 0) * (safeNum(r[idx[dashboardHeaders.unitCost]]) || 0);
    const units90 = safeNum(r[idx[dashboardHeaders.unitsSold90]]) || 0; 

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
    sheet.getRange(1, 1, out.length, out[0].length).setValues(out);
    applyHeaderStyle(sheet.getRange(1, 1, 1, out[0].length));
  }
  if (out.length > 1) {
    sheet.getRange(2, 3, out.length - 1, 1).setNumberFormat("$#,##0.00");
  }
}

function logElasticitySnapshot(ss, rows, idx) {
  const sheet = ss.getSheetByName(VDM_CONFIG.TABS.ELASTICITY) || ss.insertSheet(VDM_CONFIG.TABS.ELASTICITY);
  const date = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
  const dashboardHeaders = getDashboardReportHeaders(idx);
  
  const snapshot = rows.map(r => [
    date,
    r[idx["SKU ANCHOR KEY"]],
    r[idx["VDM MARKDOWN DEPTH %"]],
    r[idx[dashboardHeaders.simulatedNetPrice]],
    r[idx[dashboardHeaders.unitsSold90]]
  ]);
  
  const snapshotHeaders = ["Snapshot Date", "SKU", "Markdown Depth", "Price", "Velocity"];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(snapshotHeaders);
    applyHeaderStyle(sheet.getRange(1, 1, 1, snapshotHeaders.length));
  }
  if (snapshot.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, snapshot.length, snapshot[0].length).setValues(snapshot);
  }
}

/**
 * State Recovery Pattern: Recovers dashboard data from the sheet if needed for modular reporting.
 */
function recoverDashboardState() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName(VDM_CONFIG.TABS.DASHBOARD);
  const data = (dashSheet && dashSheet.getLastRow() > 0) ? dashSheet.getDataRange().getValues() : [];
  return { headers: (data[0] || []), rows: (data.slice(1) || []) };
}

function executeFlexibleRefreshProcess() {
  const ui = SpreadsheetApp.getUi();
  try {
    ui.showModelessDialog(HtmlService.createHtmlOutput("<b>Executing Full VDM System Sync...</b>"), "System Status");
    runDataIngestion();
    const dashboardState = executeDashboardRefresh();
    const reportState = generateAllReports(dashboardState);
    if (!reportState.success) {
      const failedList = reportState.failedReports.length ? reportState.failedReports.join(", ") : "Unknown";
      const details = reportState.errorDetails ? `\n\nDetails:\n${reportState.errorDetails}` : "";
      ui.alert("Partial Failure", `One or more reports failed to generate.\nFailed Reports: ${failedList}${details}`, ui.ButtonSet.OK);
      return;
    }
    const summaryTelemetry = reportState && reportState.summaryTelemetry ? reportState.summaryTelemetry : {};
    const stats = dashboardState.stats || {};
    const clearanceRatio = safeNum(summaryTelemetry.clearanceRatio) || 0;
    const heroShrinkage = safeNum(summaryTelemetry.heroShrinkage) || 0;
    const projectedRevenueImpact90 = safeNum(summaryTelemetry.projectedRevenueImpact90) || 0;
    const clearanceHealthLine = (summaryTelemetry.clearanceCapWarn ? "🚨 ALERT" : "✓ Within Cap") +
      ` (${summaryTelemetry.clearanceCount || 0}/${summaryTelemetry.totalRows || 0}; ${(clearanceRatio * 100).toFixed(2)}%)`;
    const heroHealthLine = (summaryTelemetry.heroDepletionAlert ? "🚨 ALERT" : "✓ Stable") +
      ` (Baseline ${summaryTelemetry.baselineHeroCount || 0} → Target ${summaryTelemetry.targetHeroCount || 0}; ${(heroShrinkage * 100).toFixed(2)}%)`;
    ui.alert(
      "🎉 VDM Refresh Complete! (Engine Version: v" + VDM_CONFIG.VERSION + ")\n\n" +
      "• Processed Active SKUs: " + (stats.total || 0) + "\n" +
      "• Queue 1A Missing Cost Errors: " + (stats.missingCost || 0) + "\n" +
      "• Queue 1A Negative Margin Audits: " + (stats.negativeMarginAudits || 0) + "\n" +
      "• Queue 1B Margin Floor Violators (<20%): " + (stats.blockedByMargin || 0) + "\n" +
      "• B2B Reserve Holds: " + (stats.b2bHolds || 0) + "\n" +
      "• Unmapped SHARED Physical Stock: " + (stats.missingInventory || 0) + "\n" +
      "• Fulfillment Type Fallbacks: " + (stats.fulfillmentFallbackCount || 0) + "\n" +
      "• Clearance Health (65% Markdown Cap): " + clearanceHealthLine + "\n" +
      "• Hero Pool Shrinkage: " + heroHealthLine + "\n" +
      "• Projected 90-Day Revenue Impact: $" + projectedRevenueImpact90.toFixed(2)
    );
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
    const folder = DriveApp.getFolderById(getOperationalFolderId());
    ingestEEI(folder, VDM_CONFIG.SOURCE_FILES.EEI_USA, VDM_CONFIG.TABS.RAW_EEI_USA, ss);
    ingestEEI(folder, VDM_CONFIG.SOURCE_FILES.EEI_WEB, VDM_CONFIG.TABS.RAW_EEI_WEB, ss);
    ui.alert("Inventory Snapshot Sync Complete.");
  } catch (e) { ui.alert("Sync Failed: " + e.message); }
}

function workflowIngestMetadataOnly() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const folder = DriveApp.getFolderById(getOperationalFolderId());
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
    const reportState = generateAllReports(dashboardState);
    if (!reportState.success) {
      const failedList = reportState.failedReports.length ? reportState.failedReports.join(", ") : "Unknown";
      const details = reportState.errorDetails ? `\n\nDetails:\n${reportState.errorDetails}` : "";
      ui.alert("Partial Failure", `Matrix recalculation completed, but one or more reports failed.\nFailed Reports: ${failedList}${details}`, ui.ButtonSet.OK);
      return;
    }
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
    ui.alert(getSyncAuditRefreshSuccessMessage());
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