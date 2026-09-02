/**
 * MODULE 3: MATRIX ENGINE (VDM v3.1-LEAN)
 */

const DASHBOARD_HEADERS = [
  "SKU Anchor Key",
  "Gatekeeper Status",
  "Gatekeeper Code",
  "Fulfillment Tag",
  "Resolved Cost Base",
  "Live Storefront Price",
  "Live Compare MSRP",
  "Active Storefront Markdown Depth %",
  "Current Gross Margin %",
  "Raw 90D Retail Velocity",
  "Retail Velocity Score Component",
  "Margin Score Component",
  "Retail Stock Score Component",
  "Total Composite Score",
  "Target Strategic Tier",
  "Tier Code",
  "VDM Markdown Depth %",
  "Operational Network Stock",
  "EEI Web Warehouse On Hand Stock",
  "EEI USA Warehouse Stock",
  "Live Storefront Shopify Qty",
  "WMS Sync Drift",
  "New Proposed Storefront Price",
  "Simulated Checkout Net Price",
  "Final Simulated Stacked Margin %",
  "Profit Guardrail Status Alert",
  "Guardrail Code",
  "Pricing Migration Status",
  "Retail Price Shift ($)",
  "Net Margin Change %",
  "Action Queue",
  "Queue Code"
];

function scoreMarginComponent(curMargin) {
  if (curMargin >= 0.65) return 3;
  if (curMargin >= 0.50) return 2;
  if (curMargin >= 0.35) return 1;
  return 0;
}

function scoreStockComponent(fulfillment, usaStock, shopifyQty, units90) {
  if (fulfillment === "WEBONLY") return 2;
  const operationalStock = (safeNum(usaStock) || 0) + (safeNum(shopifyQty) || 0);
  const velocity = safeNum(units90) || 0;
  const dailyVelocity = velocity > 0 ? velocity / 90 : 0;
  const dos = dailyVelocity > 0 ? operationalStock / dailyVelocity : (operationalStock > 0 ? 999 : 0);
  if (dos <= 30) return 3;
  if (dos <= 120) return 2;
  if (dos <= 180) return 1;
  return 0;
}

function executeDashboardRefresh() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const control = loadControlPanelConfig(ss);

  const shopifySheet = ss.getSheetByName(VDM_CONFIG.TABS.SHOPIFY_EXPORT);
  const usaSheet = ss.getSheetByName(VDM_CONFIG.TABS.EEI_USA);
  const webSheet = ss.getSheetByName(VDM_CONFIG.TABS.EEI_WEB);
  const salesSheet = ss.getSheetByName(VDM_CONFIG.TABS.SALES_90D);
  const resolvedCostSheet = ss.getSheetByName(VDM_CONFIG.TABS.RESOLVED_COST);
  const costLedgerSheet = ss.getSheetByName(VDM_CONFIG.TABS.COST_LEDGER);

  if (!shopifySheet || !usaSheet || !webSheet || !salesSheet || !resolvedCostSheet || !costLedgerSheet) {
    throw new Error("One or more required tabs are missing for VDM execution.");
  }

  const shopifyData = shopifySheet.getDataRange().getValues();
  if (shopifyData.length < 2) {
    writeTableToSheet_(ss, VDM_CONFIG.TABS.DASHBOARD, DASHBOARD_HEADERS, []);
    writeActionHubFromRecords_(ss, []);
    writeTierSummaryFromRecords_(ss, []);
    writeShopifyOutputFromRecords_(ss, []);
    writeSyncAuditFromRecords_(ss, []);
    writeMasterLedgerFromRecords_(ss, []);
    writeMatrixBackup_(ss, DASHBOARD_HEADERS, []);
    writeTableToSheet_(ss, VDM_CONFIG.TABS.BI_FEED, DASHBOARD_HEADERS, []);
    return { headers: DASHBOARD_HEADERS, rows: [], stats: { totalActiveSkus: 0 } };
  }

  const sIdx = getHeaderMap(shopifyData[0]);
  const skuHeader = getFirstAvailableHeader(sIdx, ["SKU_ANCHOR", "VARIANT SKU", "SKU"]);
  const statusHeader = getFirstAvailableHeader(sIdx, ["STATUS"]);
  const vendorHeader = findFirstAvailableHeader(sIdx, ["VENDOR"]);
  const titleHeader = findFirstAvailableHeader(sIdx, ["TITLE", "PRODUCT TITLE"]);
  const handleHeader = findFirstAvailableHeader(sIdx, ["HANDLE"]);
  const option1Header = findFirstAvailableHeader(sIdx, ["OPTION1 VALUE", "OPTION1"]);
  const priceHeader = getFirstAvailableHeader(sIdx, ["VARIANT PRICE", "PRICE"]);
  const compareHeader = findFirstAvailableHeader(sIdx, ["VARIANT COMPARE AT PRICE", "COMPARE AT PRICE"]);
  const qtyHeader = findFirstAvailableHeader(sIdx, ["VARIANT INVENTORY QTY", "INVENTORY QTY", "QTY", "QUANTITY"]);
  const fulfillmentHeader = findFirstAvailableHeader(sIdx, ["FULFILLMENT TYPE", "FULFILLMENT TAG"]);

  const salesMap = readSalesMap_(salesSheet);
  const resolvedCostMap = readResolvedCostMap_(resolvedCostSheet);
  const usaAgg = readWarehouseAggregate_(usaSheet, true);
  const webAgg = readWarehouseAggregate_(webSheet, false);
  const leadTimeMap = readLeadTimeMap_(costLedgerSheet, control);

  const activeRows = [];
  for (let i = 1; i < shopifyData.length; i++) {
    const row = shopifyData[i];
    const status = safeStr(row[sIdx[statusHeader]]).toLowerCase();
    if (status !== "active") continue;
    const sku = safeStr(row[sIdx[skuHeader]]).toUpperCase();
    if (!sku) continue;
    activeRows.push(row);
  }

  const percentileBase = activeRows
    .map(row => {
      const sku = safeStr(row[sIdx[skuHeader]]).toUpperCase();
      return safeNum(salesMap.get(sku)) || 0;
    })
    .filter(units => units >= 2)
    .sort((a, b) => a - b);

  const records = activeRows.map(row => {
    const sku = safeStr(row[sIdx[skuHeader]]).toUpperCase();
    const vendor = vendorHeader ? safeStr(row[sIdx[vendorHeader]]) : "";
    const vendorUpper = vendor.toUpperCase();
    const title = titleHeader ? safeStr(row[sIdx[titleHeader]]) : "";
    const handle = handleHeader ? safeStr(row[sIdx[handleHeader]]) : "";
    const option1 = option1Header ? safeStr(row[sIdx[option1Header]]) : "";

    const livePrice = safeNum(row[sIdx[priceHeader]]) || 0;
    const compareRaw = compareHeader ? safeNum(row[sIdx[compareHeader]]) : null;
    const resolvedMsrp = compareRaw !== null && compareRaw > livePrice ? compareRaw : livePrice;
    const activeMarkdown = resolvedMsrp > 0 ? Math.max(0, (resolvedMsrp - livePrice) / resolvedMsrp) : 0;

    const resolvedCost = safeNum(resolvedCostMap.get(sku)) || 0;
    const shopifyQty = qtyHeader ? (safeNum(row[sIdx[qtyHeader]]) || 0) : 0;
    const usaStock = safeNum(usaAgg.stock.get(sku)) || 0;
    const webStock = safeNum(webAgg.stock.get(sku)) || 0;
    const usa30dSales = safeNum(usaAgg.sales30d.get(sku)) || 0;
    const operationalStock = usaStock + webStock;
    const wmsSyncDrift = shopifyQty - webStock;
    const driftRatio = Math.abs(shopifyQty - webStock) / Math.max(Math.abs(shopifyQty), Math.abs(webStock), 1);

    const units90 = safeNum(salesMap.get(sku)) || 0;
    const currentMargin = livePrice > 0 ? (livePrice - resolvedCost) / livePrice : (resolvedCost > 0 ? -1 : 0);

    const ingestedFulfillment = fulfillmentHeader ? row[sIdx[fulfillmentHeader]] : "";
    const fulfillmentTag = resolveFulfillmentType(sku, ingestedFulfillment, control.virtualSkuPrefixes);

    let gatekeeperCode = GATEKEEPER_CODES.NONE;
    let gatekeeperStatus = "NONE";
    let targetTier = "";
    let markdownDepth = 0;
    let queueTags = [];

    if (resolvedCost <= 0 && !control.gwpSkus.has(sku)) {
      gatekeeperCode = GATEKEEPER_CODES.MISSING_COST;
      gatekeeperStatus = "GK_MISSING_COST: Missing Cost Base";
      targetTier = "Missing Cost Base";
      queueTags.push("Queue 1A: Missing Cost Base");
      markdownDepth = 0;
    } else if (control.gwpSkus.has(sku)) {
      gatekeeperCode = GATEKEEPER_CODES.GWP;
      gatekeeperStatus = "GK_GWP: Active GWP (0% Hold)";
      targetTier = "Active GWP (0% Hold)";
      markdownDepth = 0;
    } else if (control.newLaunchSkus.has(sku)) {
      gatekeeperCode = GATEKEEPER_CODES.LAUNCH;
      gatekeeperStatus = "GK_LAUNCH: New Launch (0% Hold)";
      targetTier = "New Launch (0% Hold)";
      markdownDepth = 0;
    } else if (control.enforceOosReversion && shopifyQty <= 0) {
      gatekeeperCode = GATEKEEPER_CODES.OOS;
      gatekeeperStatus = "GK_OOS: OOS MSRP Reversion (0% Hold)";
      targetTier = "OOS MSRP Reversion (0% Hold)";
      markdownDepth = 0;
    } else if (control.mapVendors.has(vendorUpper)) {
      gatekeeperCode = GATEKEEPER_CODES.MAP;
      gatekeeperStatus = "GK_MAP: 3P MAP Protection (0% Hold)";
      targetTier = "3P MAP Protection (0% Hold)";
      markdownDepth = 0;
    } else if (currentMargin < 0) {
      gatekeeperCode = GATEKEEPER_CODES.NEG_MARGIN;
      gatekeeperStatus = "GK_NEG_MARGIN: Negative Base Margin";
      targetTier = "Negative Base Margin";
      markdownDepth = 0;
      queueTags.push("Queue 1A: Negative Base Margin");
    }

    let velocityScore = 0;
    let marginScore = 0;
    let stockScore = 0;
    let totalScore = 0;
    let dailyVelocity = units90 > 0 ? units90 / 90 : 0;
    let daysOfSupply = dailyVelocity > 0 ? operationalStock / dailyVelocity : (operationalStock > 0 ? 999 : 0);

    if (gatekeeperCode === GATEKEEPER_CODES.NONE) {
      if (units90 === 1) {
        velocityScore = 1;
      } else if (units90 >= 2) {
        const pct = getPercentileRankInc(percentileBase, units90);
        velocityScore = pct >= 0.80 ? 4 : (pct >= 0.55 ? 3 : 2);
      }
      marginScore = scoreMarginComponent(currentMargin);
      stockScore = scoreStockComponent(fulfillmentTag, usaStock, webStock, units90);
      totalScore = velocityScore + marginScore + stockScore;

      if (totalScore === 10) {
        targetTier = "Top Hero";
        markdownDepth = 0;
      } else if (totalScore >= 8) {
        targetTier = "Signature Hero";
        markdownDepth = 0.30;
      } else if (totalScore >= 6) {
        targetTier = "Proven Performer";
        markdownDepth = 0.40;
      } else if (totalScore >= 4) {
        targetTier = "Accelerator";
        markdownDepth = 0.50;
      } else if (fulfillmentTag === "WEBONLY") {
        if (units90 === 0) {
          targetTier = "Clearance / Archive";
          markdownDepth = 0.65;
        } else {
          targetTier = "Digital Review";
          markdownDepth = 0.50;
        }
        queueTags.push("Queue 2: WEBONLY Digital Review");
      } else {
        targetTier = "Clearance / Archive";
        markdownDepth = 0.65;
        queueTags.push("Queue 3: SHARED Clearance");
      }
    }

    let proposedPrice = livePrice;
    let proposedCompareAt = resolvedMsrp;
    let b2bHold = false;

    if (gatekeeperCode === GATEKEEPER_CODES.LAUNCH) {
      proposedPrice = resolvedMsrp;
      proposedCompareAt = resolvedMsrp;
    } else if (gatekeeperCode === GATEKEEPER_CODES.OOS) {
      proposedPrice = resolvedMsrp;
      proposedCompareAt = resolvedMsrp;
    } else if (gatekeeperCode !== GATEKEEPER_CODES.NONE) {
      proposedPrice = livePrice;
      proposedCompareAt = resolvedMsrp;
    } else {
      proposedPrice = resolvedMsrp * (1 - markdownDepth);
      proposedCompareAt = markdownDepth === 0 ? proposedPrice : resolvedMsrp;

      b2bHold = fulfillmentTag === "SHARED" &&
        markdownDepth >= 0.5 &&
        usaStock >= control.b2bReserveMinQty &&
        usa30dSales > 0;

      if (b2bHold) {
        targetTier = "B2B Protection Hold";
        markdownDepth = 0;
        proposedPrice = livePrice;
        proposedCompareAt = resolvedMsrp;
        queueTags = queueTags.filter(tag => tag.indexOf("Queue 3") !== 0);
        queueTags.push("B2B Protection Hold");
      }
    }

    let simulatedNetPrice = proposedPrice * (1 - control.affiliateRate);
    let finalStackedMargin = simulatedNetPrice > 0 ? (simulatedNetPrice - resolvedCost) / simulatedNetPrice : 0;

    let guardrailStatus = "SAFE";
    let guardrailCode = GUARDRAIL_CODES.SAFE;
    if (gatekeeperCode === GATEKEEPER_CODES.MISSING_COST) {
      guardrailStatus = "❌ BLOCKED (Missing Cost Base)";
      guardrailCode = GATEKEEPER_CODES.MISSING_COST;
    } else if (gatekeeperCode === GATEKEEPER_CODES.NEG_MARGIN) {
      guardrailStatus = "❌ BLOCKED (Negative Base Margin)";
      guardrailCode = GATEKEEPER_CODES.NEG_MARGIN;
    } else if (b2bHold) {
      guardrailStatus = "⚠️ B2B HOLD";
      guardrailCode = GUARDRAIL_CODES.WARN_B2B_HOLD;
    }

    if (gatekeeperCode === GATEKEEPER_CODES.NONE && !b2bHold && finalStackedMargin < VDM_CONFIG.FLOORS.STACKED_MARGIN_MIN) {
      guardrailStatus = "❌ BLOCKED (Margin Floor Violator)";
      guardrailCode = GUARDRAIL_CODES.BLOCK_MARGIN_FLOOR;
      queueTags = queueTags.filter(tag => tag.indexOf("Queue 2") !== 0 && tag.indexOf("Queue 3") !== 0);
      queueTags.push("Queue 1B: Margin Floor Violator");
      proposedPrice = livePrice;
      proposedCompareAt = livePrice;
      markdownDepth = 0;
      simulatedNetPrice = proposedPrice * (1 - control.affiliateRate);
      finalStackedMargin = simulatedNetPrice > 0 ? (simulatedNetPrice - resolvedCost) / simulatedNetPrice : 0;
    }

    if (driftRatio >= VDM_CONFIG.FLOORS.DRIFT_ALERT_RATIO) {
      queueTags.push("Queue 1A: Inventory Sync Drift Alert");
    }

    const targetLeadTime = leadTimeMap.has(sku)
      ? leadTimeMap.get(sku)
      : control.defaultDomesticLeadTime;

    let reorderStatus = "NO_REORDER_SIGNAL";
    let suggestedReorderQty = 0;
    if (totalScore <= 3 && fulfillmentTag === "SHARED") {
      reorderStatus = "OK (Clearance)";
    } else if (gatekeeperCode === GATEKEEPER_CODES.NONE && units90 > 0) {
      if (daysOfSupply <= targetLeadTime) {
        reorderStatus = "REORDER_ALERT";
        suggestedReorderQty = Math.max(0, Math.ceil((targetLeadTime * dailyVelocity) - operationalStock));
      }
    }

    if (reorderStatus === "REORDER_ALERT") {
      queueTags.push("REORDER_ALERT");
    }

    queueTags = Array.from(new Set(queueTags));

    const pricingMigrationStatus = proposedPrice > livePrice
      ? "📈 Price Recovery/Lift"
      : (proposedPrice < livePrice ? "🚨 Deepen Discount" : "✓ Price Hold");

    return {
      sku,
      gatekeeperCode,
      gatekeeperStatus,
      fulfillmentTag,
      resolvedCost,
      livePrice,
      resolvedMsrp,
      activeMarkdown,
      currentMargin,
      units90,
      velocityScore,
      marginScore,
      stockScore,
      totalScore,
      targetTier,
      tierCode: resolveTierCode_(targetTier, gatekeeperCode),
      markdownDepth,
      operationalStock,
      webStock,
      usaStock,
      shopifyQty,
      wmsSyncDrift,
      proposedPrice,
      proposedCompareAt,
      simulatedNetPrice,
      finalStackedMargin,
      guardrailStatus,
      guardrailCode,
      pricingMigrationStatus,
      priceShift: proposedPrice - livePrice,
      marginChange: finalStackedMargin - currentMargin,
      actionQueue: queueTags.join(" | "),
      queueCode: resolveQueueCode_(queueTags, b2bHold),
      title,
      handle,
      option1,
      vendor,
      daysOfSupply,
      reorderStatus,
      suggestedReorderQty
    };
  });

  const dashboardRows = records.map(record => [
    record.sku,
    record.gatekeeperStatus,
    record.gatekeeperCode,
    record.fulfillmentTag,
    record.resolvedCost,
    record.livePrice,
    record.resolvedMsrp,
    record.activeMarkdown,
    record.currentMargin,
    record.units90,
    record.velocityScore,
    record.marginScore,
    record.stockScore,
    record.totalScore,
    record.targetTier,
    record.tierCode,
    record.markdownDepth,
    record.operationalStock,
    record.webStock,
    record.usaStock,
    record.shopifyQty,
    record.wmsSyncDrift,
    record.proposedPrice,
    record.simulatedNetPrice,
    record.finalStackedMargin,
    record.guardrailStatus,
    record.guardrailCode,
    record.pricingMigrationStatus,
    record.priceShift,
    record.marginChange,
    record.actionQueue,
    record.queueCode
  ]);

  writeTableToSheet_(ss, VDM_CONFIG.TABS.DASHBOARD, DASHBOARD_HEADERS, dashboardRows);
  writeActionHubFromRecords_(ss, records);
  writeTierSummaryFromRecords_(ss, records);
  writeShopifyOutputFromRecords_(ss, records);
  writeSyncAuditFromRecords_(ss, records);
  writeMasterLedgerFromRecords_(ss, records);
  writeMatrixBackup_(ss, DASHBOARD_HEADERS, dashboardRows);
  writeTableToSheet_(ss, VDM_CONFIG.TABS.BI_FEED, DASHBOARD_HEADERS, dashboardRows);

  return {
    headers: DASHBOARD_HEADERS,
    rows: dashboardRows,
    stats: {
      totalActiveSkus: records.length,
      gatekeeperCount: records.filter(r => r.gatekeeperCode !== GATEKEEPER_CODES.NONE).length,
      reorderAlertCount: records.filter(r => r.reorderStatus === "REORDER_ALERT").length
    }
  };
}

function writeActionHubFromRecords_(ss, records) {
  const width = 12;
  const rows = [];

  rows.push(["Section A: Catalog Data Triage", "", "", "", "", "", "", "", "", "", "", ""]);
  rows.push([
    "SKU Anchor Key",
    "Product Title",
    "Issue Type",
    "Unit Cost",
    "Variant Price",
    "Shopify Qty",
    "Total Network Stock",
    "Action Required",
    "",
    "",
    "",
    ""
  ]);

  records.forEach(r => {
    if (r.actionQueue.indexOf("Queue 1A") === -1) return;
    const issue = r.actionQueue.split("|").map(v => safeStr(v)).find(v => v.indexOf("Queue 1A") === 0) || "Queue 1A";
    rows.push([
      r.sku,
      r.title,
      issue,
      r.resolvedCost,
      r.livePrice,
      r.shopifyQty,
      r.operationalStock,
      "Resolve catalog data exception before pricing publish.",
      "",
      "",
      "",
      ""
    ]);
  });

  rows.push(["", "", "", "", "", "", "", "", "", "", "", ""]);
  rows.push(["Section B: Merchandising Sign-Off & Reorder Hub", "", "", "", "", "", "", "", "", "", "", ""]);
  rows.push([
    "Handle",
    "Title",
    "Option1 Value",
    "Variant SKU",
    "Current Price",
    "Proposed Price",
    "Proposed Compare At",
    "Stacked Margin %",
    "Days of Supply",
    "Reorder Status",
    "Suggested Reorder Qty",
    "Queue Category"
  ]);

  records.forEach(r => {
    const inSectionB =
      r.actionQueue.indexOf("Queue 1B") !== -1 ||
      r.actionQueue.indexOf("Queue 2") !== -1 ||
      r.actionQueue.indexOf("Queue 3") !== -1 ||
      r.actionQueue.indexOf("B2B Protection Hold") !== -1 ||
      r.reorderStatus === "REORDER_ALERT";
    if (!inSectionB) return;

    const queueCategory = r.reorderStatus === "REORDER_ALERT"
      ? "REORDER_ALERT"
      : (r.actionQueue.split("|").map(v => safeStr(v)).find(v => /^Queue\s[123]/.test(v) || v === "B2B Protection Hold") || "");

    rows.push([
      r.handle,
      r.title,
      r.option1,
      r.sku,
      r.livePrice,
      r.proposedPrice,
      r.proposedCompareAt,
      r.finalStackedMargin,
      r.daysOfSupply,
      r.reorderStatus,
      r.suggestedReorderQty,
      queueCategory
    ]);
  });

  const normalizedRows = rows.map(row => {
    if (row.length !== width) {
      throw new Error(`Action Hub row width mismatch. Expected ${width}, got ${row.length}.`);
    }
    return row;
  });

  writeTableToSheet_(ss, VDM_CONFIG.TABS.ACTION_HUB, new Array(width).fill(""), normalizedRows, true);
}

function writeTierSummaryFromRecords_(ss, records) {
  const headers = [
    "Strategic Tier",
    "VDM Proposed SKU Count",
    "House Brand Count (Gläs / Glastoy)",
    "Avg. Stacked Checkout Margin %"
  ];

  const tierOrder = [
    "Top Hero",
    "Signature Hero",
    "Proven Performer",
    "Accelerator",
    "Digital Review",
    "Clearance / Archive"
  ];

  const rows = tierOrder.map(tier => {
    const members = records.filter(r => r.targetTier === tier);
    const houseBrandCount = members.filter(r => isHouseBrandVendor_(r.vendor)).length;
    const avgMargin = members.length
      ? members.reduce((sum, r) => sum + (safeNum(r.finalStackedMargin) || 0), 0) / members.length
      : 0;

    return [tier, members.length, houseBrandCount, avgMargin];
  });

  writeTableToSheet_(ss, VDM_CONFIG.TABS.TIER_SUMMARY, headers, rows);
}

function writeShopifyOutputFromRecords_(ss, records) {
  const headers = [
    "Handle",
    "Title",
    "Option1 Value",
    "Variant SKU",
    "Variant Price",
    "Variant Compare At Price"
  ];

  const rows = records.map(r => {
    return [
      r.handle,
      r.title,
      r.option1,
      r.sku,
      r.proposedPrice,
      r.proposedCompareAt
    ];
  });

  writeTableToSheet_(ss, VDM_CONFIG.TABS.SHOPIFY_OUTPUT, headers, rows);
}

function writeSyncAuditFromRecords_(ss, records) {
  const headers = [
    "SKU Key",
    "Handle",
    "Action",
    "Strategy Tier",
    "VDM Markdown %",
    "Old Price",
    "New Price",
    "Old MSRP",
    "New MSRP",
    "Base Price",
    "Guardrail"
  ];
  const rows = records.map(r => [
    r.sku,
    r.handle,
    r.pricingMigrationStatus,
    r.targetTier,
    r.markdownDepth,
    r.livePrice,
    r.proposedPrice,
    r.resolvedMsrp,
    r.proposedCompareAt,
    r.livePrice,
    r.guardrailStatus
  ]);
  writeTableToSheet_(ss, VDM_CONFIG.TABS.SYNC_AUDIT, headers, rows);
}

function writeMasterLedgerFromRecords_(ss, records) {
  const headers = [
    "SKU Key",
    "Handle",
    "Fulfillment",
    "Gatekeeper",
    "Migration Status",
    "Tier",
    "Target Mkdn %",
    "Old Price",
    "New Price",
    "Price Shift ($)",
    "Procurement Cost",
    "Checkout Price",
    "Stacked Margin %",
    "Guardrail",
    "Margin Shift %"
  ];
  const rows = records.map(r => [
    r.sku,
    r.handle,
    r.fulfillmentTag,
    r.gatekeeperCode,
    r.pricingMigrationStatus,
    r.targetTier,
    r.markdownDepth,
    r.livePrice,
    r.proposedPrice,
    r.priceShift,
    r.resolvedCost,
    r.simulatedNetPrice,
    r.finalStackedMargin,
    r.guardrailStatus,
    r.marginChange
  ]);
  writeTableToSheet_(ss, VDM_CONFIG.TABS.MASTER_LEDGER, headers, rows);
}

function writeMatrixBackup_(ss, headers, rows) {
  writeTableToSheet_(ss, VDM_CONFIG.TABS.BACKUP_MATRIX_DATA, headers, rows);
}

function resolveTierCode_(targetTier, gatekeeperCode) {
  if (gatekeeperCode !== GATEKEEPER_CODES.NONE || targetTier === "B2B Protection Hold") return "TIER_00_GATEKEEPER";
  if (targetTier === "Top Hero") return "TIER_10_TOP_HERO";
  if (targetTier === "Signature Hero") return "TIER_08_SIG_HERO";
  if (targetTier === "Proven Performer") return "TIER_06_PROVEN_PERFORMER";
  if (targetTier === "Accelerator" || targetTier === "Digital Review") return "TIER_04_ACCELERATOR";
  if (targetTier === "Clearance / Archive") return "TIER_00_CLEARANCE";
  return "TIER_UNKNOWN";
}

function resolveQueueCode_(queueTags, b2bHold) {
  if (b2bHold) return QUEUE_CODES.NONE;
  const tags = Array.isArray(queueTags) ? queueTags : [];
  if (tags.some(tag => tag.indexOf("Queue 1A: Missing Cost Base") === 0)) return QUEUE_CODES.QUEUE_1A;
  if (tags.some(tag => tag.indexOf("Queue 1A: Negative Base Margin") === 0)) return QUEUE_CODES.QUEUE_1A;
  if (tags.some(tag => tag.indexOf("Queue 1B: Margin Floor Violator") === 0)) return QUEUE_CODES.QUEUE_1B;
  if (tags.some(tag => tag.indexOf("Queue 2") === 0)) return QUEUE_CODES.QUEUE_2;
  if (tags.some(tag => tag.indexOf("Queue 3") === 0)) return QUEUE_CODES.QUEUE_3;
  if (tags.some(tag => tag.indexOf("REORDER_ALERT") === 0)) return QUEUE_CODES.REORDER_ALERT;
  return QUEUE_CODES.NONE;
}

function writeActionHubFromDashboardState(dashboardState) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const records = inflateRecordsFromDashboard_(dashboardState);
  writeActionHubFromRecords_(ss, records);
}

function writeTierSummaryFromDashboardState(dashboardState) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const records = inflateRecordsFromDashboard_(dashboardState);
  writeTierSummaryFromRecords_(ss, records);
}

function writeSyncAuditFromDashboardState(dashboardState) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const records = inflateRecordsFromDashboard_(dashboardState);
  writeSyncAuditFromRecords_(ss, records);
}

function writeMasterLedgerFromDashboardState(dashboardState) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const records = inflateRecordsFromDashboard_(dashboardState);
  writeMasterLedgerFromRecords_(ss, records);
}

function inflateRecordsFromDashboard_(dashboardState) {
  if (!dashboardState || !Array.isArray(dashboardState.headers) || !Array.isArray(dashboardState.rows)) {
    throw new Error("Dashboard state is missing or invalid.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shopifySheet = ss.getSheetByName(VDM_CONFIG.TABS.SHOPIFY_EXPORT);
  const shopifyData = shopifySheet ? shopifySheet.getDataRange().getValues() : [];
  const sMap = new Map();
  if (shopifyData.length > 1) {
    const idx = getHeaderMap(shopifyData[0]);
    const skuHeader = getFirstAvailableHeader(idx, ["SKU_ANCHOR", "VARIANT SKU", "SKU"]);
    const titleHeader = findFirstAvailableHeader(idx, ["TITLE", "PRODUCT TITLE"]);
    const handleHeader = findFirstAvailableHeader(idx, ["HANDLE"]);
    const option1Header = findFirstAvailableHeader(idx, ["OPTION1 VALUE", "OPTION1"]);
    const vendorHeader = findFirstAvailableHeader(idx, ["VENDOR"]);
    shopifyData.slice(1).forEach(row => {
      const sku = safeStr(row[idx[skuHeader]]).toUpperCase();
      if (!sku) return;
      sMap.set(sku, {
        title: titleHeader ? safeStr(row[idx[titleHeader]]) : "",
        handle: handleHeader ? safeStr(row[idx[handleHeader]]) : "",
        option1: option1Header ? safeStr(row[idx[option1Header]]) : "",
        vendor: vendorHeader ? safeStr(row[idx[vendorHeader]]) : ""
      });
    });
  }

  const hMap = getHeaderMap(dashboardState.headers);
  const getValue = (row, name) => {
    const i = hMap[safeStr(name).toUpperCase()];
    return i === undefined ? null : row[i];
  };

  return dashboardState.rows.map(row => {
    const sku = safeStr(getValue(row, "SKU Anchor Key")).toUpperCase();
    const meta = sMap.get(sku) || {};
    const proposedPrice = safeNum(getValue(row, "New Proposed Storefront Price")) || 0;
    const markdown = safeNum(getValue(row, "VDM Markdown Depth %")) || 0;
    const resolvedMsrp = safeNum(getValue(row, "Live Compare MSRP")) || 0;
    const targetTier = safeStr(getValue(row, "Target Strategic Tier"));
    const gatekeeperCode = safeStr(getValue(row, "Gatekeeper Code")) || GATEKEEPER_CODES.NONE;
    const preserveCompareAtOnZeroMarkdown = targetTier === "B2B Protection Hold" ||
      (gatekeeperCode !== GATEKEEPER_CODES.NONE && gatekeeperCode !== GATEKEEPER_CODES.OOS);

    const units90 = safeNum(getValue(row, "Raw 90D Retail Velocity")) || 0;
    const operationalStock = safeNum(getValue(row, "Operational Network Stock")) || 0;
    const dailyVelocity = units90 > 0 ? units90 / 90 : 0;
    const daysOfSupply = dailyVelocity > 0 ? operationalStock / dailyVelocity : (operationalStock > 0 ? 999 : 0);
    const totalScore = safeNum(getValue(row, "Total Composite Score")) || 0;
    const fulfillmentTag = safeStr(getValue(row, "Fulfillment Tag"));
    const actionQueue = safeStr(getValue(row, "Action Queue"));

    return {
      sku,
      title: meta.title || "",
      handle: meta.handle || "",
      option1: meta.option1 || "",
      vendor: meta.vendor || "",
      resolvedCost: safeNum(getValue(row, "Resolved Cost Base")) || 0,
      livePrice: safeNum(getValue(row, "Live Storefront Price")) || 0,
      shopifyQty: safeNum(getValue(row, "Live Storefront Shopify Qty")) || 0,
      operationalStock,
      proposedPrice,
      proposedCompareAt: (markdown === 0 && !preserveCompareAtOnZeroMarkdown ? proposedPrice : resolvedMsrp),
      finalStackedMargin: safeNum(getValue(row, "Final Simulated Stacked Margin %")) || 0,
      daysOfSupply,
      reorderStatus: actionQueue.indexOf("REORDER_ALERT") !== -1
        ? "REORDER_ALERT"
        : (totalScore <= 3 && fulfillmentTag === "SHARED" ? "OK (Clearance)" : "NO_REORDER_SIGNAL"),
      suggestedReorderQty: 0,
      actionQueue,
      targetTier,
      gatekeeperCode,
      markdownDepth: markdown,
      pricingMigrationStatus: safeStr(getValue(row, "Pricing Migration Status")) || "✓ Price Hold",
      priceShift: safeNum(getValue(row, "Retail Price Shift ($)")) || 0,
      marginChange: safeNum(getValue(row, "Net Margin Change %")) || 0,
      guardrailStatus: safeStr(getValue(row, "Profit Guardrail Status Alert")) || "SAFE",
      simulatedNetPrice: safeNum(getValue(row, "Simulated Checkout Net Price")) || 0
    };
  });
}

function readWarehouseAggregate_(sheet, includeSales) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return { stock: new Map(), sales30d: new Map() };
  }

  const header = values[0] || [];
  const idx = getHeaderMap(header);
  const skuHeader = getFirstAvailableHeader(idx, ["SKU_ANCHOR", "ITEM CODE", "SKU"]);
  const qtyHeader = getFirstAvailableHeader(idx, ["EEI USA WAREHOUSE ON HAND STOCK", "EEI WEB WAREHOUSE ON HAND STOCK", "QTY", "QUANTITY", "ON HAND STOCK"]);
  const salesHeader = includeSales ? findFirstAvailableHeader(idx, ["SALES PAST 30 DAYS"]) : null;

  const stockMap = new Map();
  const salesMap = new Map();
  values.slice(1).forEach(row => {
    const sku = safeStr(row[idx[skuHeader]]).toUpperCase();
    if (!sku) return;
    const qty = safeNum(row[idx[qtyHeader]]) || 0;
    stockMap.set(sku, (stockMap.get(sku) || 0) + qty);
    if (salesHeader) {
      const sales = safeNum(row[idx[salesHeader]]) || 0;
      salesMap.set(sku, (salesMap.get(sku) || 0) + sales);
    }
  });

  return { stock: stockMap, sales30d: salesMap };
}

function readSalesMap_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return new Map();
  const idx = getHeaderMap(values[0]);
  const skuHeader = getFirstAvailableHeader(idx, ["SKU_ANCHOR", "PRODUCT VARIANT SKU", "VARIANT SKU", "SKU"]);
  const qtyHeader = getFirstAvailableHeader(idx, ["NET QUANTITY", "NET ITEMS SOLD", "QTY", "QUANTITY"]);

  const out = new Map();
  values.slice(1).forEach(row => {
    const sku = safeStr(row[idx[skuHeader]]).toUpperCase();
    if (!sku) return;
    out.set(sku, (out.get(sku) || 0) + (safeNum(row[idx[qtyHeader]]) || 0));
  });
  return out;
}

function readResolvedCostMap_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return new Map();
  const idx = getHeaderMap(values[0]);
  const skuHeader = getFirstAvailableHeader(idx, ["SKU ANCHOR", "SKU_ANCHOR", "SKU", "VARIANT SKU"]);
  const costHeader = getFirstAvailableHeader(idx, ["RESOLVED COST", "RESOLVED COST BASE", "UNIT COST"]);

  const out = new Map();
  values.slice(1).forEach(row => {
    const sku = safeStr(row[idx[skuHeader]]).toUpperCase();
    if (!sku) return;
    const cost = costHeader ? safeNum(row[idx[costHeader]]) : null;
    out.set(sku, cost !== null ? cost : 0);
  });
  return out;
}

function readLeadTimeMap_(sheet, control) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return new Map();
  const idx = getHeaderMap(values[0]);
  const skuHeader = findFirstAvailableHeader(idx, ["SKU_ANCHOR", "SKU ANCHOR", "SKU", "VARIANT SKU"]);
  const leadHeader = findFirstAvailableHeader(idx, ["LEAD TIME", "LEAD TIME (DAYS)", "LEAD DAYS", "LEADTIME", "SUPPLY LEAD TIME"]);
  const map = new Map();

  values.slice(1).forEach(row => {
    const sku = skuHeader ? safeStr(row[idx[skuHeader]]).toUpperCase() : "";
    if (!sku) return;
    const rawLead = leadHeader ? row[idx[leadHeader]] : (row.length > 6 ? row[6] : null);
    const numericLead = safeNum(rawLead);
    if (numericLead !== null && numericLead > 0) {
      map.set(sku, Math.round(numericLead));
      return;
    }
    const text = safeStr(rawLead).toUpperCase();
    if (text.indexOf("IMPORT") !== -1) {
      map.set(sku, control.defaultImportLeadTime);
    } else {
      map.set(sku, control.defaultDomesticLeadTime);
    }
  });

  return map;
}

function isHouseBrandVendor_(vendor) {
  const normalized = safeStr(vendor).toUpperCase();
  return VDM_CONFIG.HOUSE_BRANDS.some(brand => normalized.indexOf(brand.toUpperCase()) !== -1);
}

function writeTableToSheet_(ss, sheetName, headers, rows, headersAreData) {
  const headerRow = headersAreData ? null : headers;
  const width = headersAreData
    ? (rows[0] ? rows[0].length : 0)
    : headers.length;

  if (!headersAreData && width === 0) {
    throw new Error(`Cannot write ${sheetName}: headers are empty.`);
  }

  assertRectangularRows_(rows, width, sheetName);
  const payload = headersAreData ? rows : [headerRow].concat(rows);

  const sheet = getOrCreateSheet(sheetName, sheetName.indexOf("_") === 0);
  if (sheetName === VDM_CONFIG.TABS.BI_FEED) {
    resetBiFeedSheet(sheet);
  } else {
    sheet.clearContents();
  }
  if (payload.length === 0) return;
  sheet.getRange(1, 1, payload.length, width).setValues(payload);
}

function assertRectangularRows_(rows, width, sheetName) {
  rows.forEach((row, i) => {
    if (!Array.isArray(row) || row.length !== width) {
      throw new Error(
        `Rectangular assertion failed for ${sheetName} at row ${i + 1}: expected ${width} columns, got ${Array.isArray(row) ? row.length : "non-array"}.`
      );
    }
  });
}
