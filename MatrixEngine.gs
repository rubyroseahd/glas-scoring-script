/**
 * MODULE 3: MATRIX ENGINE
 */

function isZeroCostPermitted(gatekeeperCode) {
  return gatekeeperCode === GATEKEEPER_CODES.GWP;
}

function executeDashboardRefresh() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashSheet = getOrCreateSheet(VDM_CONFIG.TABS.DASHBOARD);
    
    // 1. Memory Load: Load all raw data into lookup objects
    const shopifyData = ss.getSheetByName(VDM_CONFIG.TABS.RAW_SHOPIFY)?.getDataRange().getValues() || [];
    const salesData = ss.getSheetByName(VDM_CONFIG.TABS.RAW_SALES)?.getDataRange().getValues() || [];
    const usaData = ss.getSheetByName(VDM_CONFIG.TABS.RAW_EEI_USA)?.getDataRange().getValues() || [];
    const webData = ss.getSheetByName(VDM_CONFIG.TABS.RAW_EEI_WEB)?.getDataRange().getValues() || [];
    const costData = ss.getSheetByName(VDM_CONFIG.TABS.MASTER_COST)?.getDataRange().getValues() || [];
    const settingsData = ss.getSheetByName(VDM_CONFIG.TABS.SETTINGS)?.getDataRange().getValues() || [];

    const sIdx = shopifyData.length > 0 ? getHeaderMap(shopifyData[0]) : {};
    const vIdx = salesData.length > 0 ? getHeaderMap(salesData[0]) : {};
    const uIdx = usaData.length > 0 ? getHeaderMap(usaData[0]) : {};
    const wIdx = webData.length > 0 ? getHeaderMap(webData[0]) : {};
    const cIdx = costData.length > 0 ? getHeaderMap(costData[0]) : {};

    const hasShopifyRows = shopifyData.length > 1;
    const hasSalesRows = salesData.length > 1;
    const hasUsaRows = usaData.length > 1;
    const hasWebRows = webData.length > 1;
    const shopifySkuHeader = hasShopifyRows ? getFirstAvailableHeader(sIdx, ["SKU_ANCHOR", "VARIANT SKU", "SKU"]) : findFirstAvailableHeader(sIdx, ["SKU_ANCHOR", "VARIANT SKU", "SKU"]);
    const salesSkuHeader = hasSalesRows ? getFirstAvailableHeader(vIdx, ["SKU_ANCHOR", "PRODUCT VARIANT SKU", "VARIANT SKU", "SKU"]) : findFirstAvailableHeader(vIdx, ["SKU_ANCHOR", "PRODUCT VARIANT SKU", "VARIANT SKU", "SKU"]);
    const usaSkuHeader = hasUsaRows ? getFirstAvailableHeader(uIdx, ["SKU_ANCHOR", "ITEM CODE", "SKU"]) : findFirstAvailableHeader(uIdx, ["SKU_ANCHOR", "ITEM CODE", "SKU"]);
    const webSkuHeader = hasWebRows ? getFirstAvailableHeader(wIdx, ["SKU_ANCHOR", "ITEM CODE", "SKU"]) : findFirstAvailableHeader(wIdx, ["SKU_ANCHOR", "ITEM CODE", "SKU"]);
    const hasCostRows = costData.length > 1;
    const costSkuHeader = hasCostRows ? getFirstAvailableHeader(cIdx, ["SKU_ANCHOR", "SKU ANCHOR", "SKU", "VARIANT SKU"]) : findFirstAvailableHeader(cIdx, ["SKU_ANCHOR", "SKU ANCHOR", "SKU", "VARIANT SKU"]);

    const salesValueHeader = hasSalesRows ? getFirstAvailableHeader(vIdx, ["NET QUANTITY", "NET ITEMS SOLD", "QTY", "QUANTITY"]) : findFirstAvailableHeader(vIdx, ["NET QUANTITY", "NET ITEMS SOLD", "QTY", "QUANTITY"]);
    const webStockHeader = hasWebRows ? getFirstAvailableHeader(wIdx, ["EEI WEB WAREHOUSE ON HAND STOCK", "QTY", "QUANTITY", "ON HAND STOCK"]) : findFirstAvailableHeader(wIdx, ["EEI WEB WAREHOUSE ON HAND STOCK", "QTY", "QUANTITY", "ON HAND STOCK"]);
    const usaStockHeader = hasUsaRows ? getFirstAvailableHeader(uIdx, ["EEI USA WAREHOUSE ON HAND STOCK", "QTY", "QUANTITY", "ON HAND STOCK"]) : findFirstAvailableHeader(uIdx, ["EEI USA WAREHOUSE ON HAND STOCK", "QTY", "QUANTITY", "ON HAND STOCK"]);
    const usaSalesHeader = findFirstAvailableHeader(uIdx, ["SALES PAST 30 DAYS"]);
    const shopifyPriceHeader = hasShopifyRows ? getFirstAvailableHeader(sIdx, ["VARIANT PRICE", "PRICE"]) : findFirstAvailableHeader(sIdx, ["VARIANT PRICE", "PRICE"]);
    const shopifyCompareHeader = findFirstAvailableHeader(sIdx, ["VARIANT COMPARE AT PRICE", "COMPARE AT PRICE"]);
    const shopifyQtyHeader = findFirstAvailableHeader(sIdx, ["VARIANT INVENTORY QTY", "INVENTORY QTY"]);
    const fulfillmentHeader = findFirstAvailableHeader(sIdx, ["FULFILLMENT TYPE"]);
    const vendorHeader = findFirstAvailableHeader(sIdx, ["VENDOR"]);
    const resolvedCostHeader = findFirstAvailableHeader(cIdx, ["RESOLVED COST"]);

    const salesMap = salesValueHeader && salesSkuHeader
      ? new Map(salesData.slice(1).map(r => [safeStr(r[vIdx[salesSkuHeader]]).toUpperCase(), safeNum(r[vIdx[salesValueHeader]])]))
      : new Map();
    const usaMap = usaSkuHeader
      ? new Map(usaData.slice(1).map(r => [safeStr(r[uIdx[usaSkuHeader]]).toUpperCase(), r]))
      : new Map();
    const webMap = webStockHeader && webSkuHeader
      ? new Map(webData.slice(1).map(r => [safeStr(r[wIdx[webSkuHeader]]).toUpperCase(), safeNum(r[wIdx[webStockHeader]])]))
      : new Map();
    const costMap = costSkuHeader
      ? new Map(costData.slice(1).map(r => [safeStr(r[cIdx[costSkuHeader]]).toUpperCase(), resolvedCostHeader ? safeNum(r[cIdx[resolvedCostHeader]]) : null]))
      : new Map();
    const baselineSheet = ss.getSheetByName(VDM_CONFIG.TABS.BASELINE);
    const baselineData = (baselineSheet && baselineSheet.getLastRow() > 1) ? baselineSheet.getDataRange().getValues() : [];
    const baselineMap = new Map();
    if (baselineData.length > 1) {
      const bIdx = getHeaderMap(baselineData[0]);
      const baselineSkuHeader = findFirstAvailableHeader(bIdx, ["SKU ANCHOR", "SKU_ANCHOR", "SKU ANCHOR KEY", "SKU", "VARIANT SKU"]);
      const baselinePriceHeader = findFirstAvailableHeader(bIdx, ["LIVE PRICE", "LIVE STOREFRONT PRICE", "BASELINE STOREFRONT PRICE", "PRICE", "VARIANT PRICE"]);
      const baselineCompareHeader = findFirstAvailableHeader(bIdx, ["COMPARE MSRP", "LIVE COMPARE MSRP", "VARIANT COMPARE AT PRICE", "COMPARE AT PRICE"]);
      const baselineBracketHeader = findFirstAvailableHeader(bIdx, ["BASELINE BRACKET", "BASELINE STOREFRONT BRACKET", "STOREFRONT BRACKET", "CURRENT EQUIVALENT STOREFRONT TIER"]);
      const baselineCostHeader = findFirstAvailableHeader(bIdx, ["RESOLVED COST BASE", "COST", "PROCUREMENT COST"]);
      baselineData.slice(1).forEach(row => {
        const skuKey = baselineSkuHeader ? safeStr(row[bIdx[baselineSkuHeader]]).toUpperCase() : "";
        if (!skuKey) return;
        baselineMap.set(skuKey, {
          price: baselinePriceHeader ? safeNum(row[bIdx[baselinePriceHeader]]) : null,
          compareMsrp: baselineCompareHeader ? safeNum(row[bIdx[baselineCompareHeader]]) : null,
          bracket: baselineBracketHeader ? safeStr(row[bIdx[baselineBracketHeader]]) : "",
          cost: baselineCostHeader ? safeNum(row[bIdx[baselineCostHeader]]) : null
        });
      });
    }
    const catalogSkuSet = new Set(
      shopifySkuHeader ? shopifyData.slice(1).map(r => safeStr(r[sIdx[shopifySkuHeader]]).toUpperCase()).filter(Boolean) : []
    );
    const unmappedPhysicalSkus = new Set(
      [...Array.from(usaMap.keys()), ...Array.from(webMap.keys())].filter(sku => !catalogSkuSet.has(sku))
    );
    
    // Load Registries
    const gwpSet = new Set(settingsData.slice(1).map(r => safeStr(r[0]).toUpperCase())); // Skip header row
    const launchSet = new Set(settingsData.slice(1).map(r => safeStr(r[1]).toUpperCase())); // Skip header row
    const mapVendors = settingsData.slice(1).map(r => safeStr(r[2]).toUpperCase()).filter(v => v); // Vendor-level MAP only
    const configuredAffiliateRate = settingsData.length > 1 ? safeNum(settingsData[1][4]) : null;
    const affiliateRate = configuredAffiliateRate !== null && configuredAffiliateRate >= 0 && configuredAffiliateRate < 1
      ? configuredAffiliateRate
      : VDM_CONFIG.AFFILIATE_RATE_DEFAULT;
    // B2B Reserve Min Qty from Settings column D; defaults to 500 if missing or invalid
    const b2bReserveMin = (settingsData.length > 1 && safeNum(settingsData[1][3]) !== null && safeNum(settingsData[1][3]) > 0) ? safeNum(settingsData[1][3]) : 500;

    // Velocity Percentile Setup
    const salesArray = Array.from(salesMap.values()).filter(v => v !== null && v > 1).sort((a,b) => a-b);
      
    // Data Health Tracking
    const stats = {
      total: 0,
      missingCost: 0,
      negativeMarginAudits: 0,
      missingInventory: unmappedPhysicalSkus.size,
      blockedByMargin: 0,
      b2bHolds: 0,
      fulfillmentFallbackCount: 0
    };

    const results = [];
    shopifyData.slice(1).forEach(row => {
      const sku = shopifySkuHeader ? safeStr(row[sIdx[shopifySkuHeader]]).toUpperCase() : ""; // Normalize SKU key for consistent map lookups
      if (!sku) return;
      const vendor = vendorHeader ? safeStr(row[sIdx[vendorHeader]]).toUpperCase() : "";
      
      // A: SKU Anchor
      // B: Gatekeeper
      let gate = "None";
      let gateCode = GATEKEEPER_CODES.NONE;
      if (gwpSet.has(sku)) { gate = "⚠️ Active GWP Promo"; gateCode = GATEKEEPER_CODES.GWP; }
      else if (launchSet.has(sku)) { gate = "New Launch"; gateCode = GATEKEEPER_CODES.NEW_LAUNCH; }
      else if (isMapVendorMatch(vendor, mapVendors)) { gate = "3rd Party MAP"; gateCode = GATEKEEPER_CODES.MAP; }

      const fulfillment = fulfillmentHeader ? (safeStr(row[sIdx[fulfillmentHeader]]).toUpperCase() || "SHARED") : "SHARED";
      if (!fulfillmentHeader || !safeStr(row[sIdx[fulfillmentHeader]])) stats.fulfillmentFallbackCount++;
      const cost = safeNum(costMap.get(sku));
      const price = safeNum(row[sIdx[shopifyPriceHeader]]);
      const rawCompare = shopifyCompareHeader ? safeNum(row[sIdx[shopifyCompareHeader]]) : null;
      
      const compareMSRP = (rawCompare === 0 || rawCompare === null) ? (price || 0) : rawCompare;
      const curMarkdown = (compareMSRP === price || compareMSRP === 0) ? 0 : (compareMSRP - (price || 0)) / compareMSRP;
      const curMargin = (price === 0 || price === null || cost === null) ? 0 : (price - cost) / price;
      
      // Velocity Score (I)
      const units90 = safeNum(salesMap.get(sku));
      let vScore = 0;
      if (units90 !== null) {
        if (units90 === 1) {
          vScore = 1;
        } else if (units90 > 1) {
          const pct = getPercentileRankInc(salesArray, units90);
          if (pct >= 0.80) vScore = 4;
          else if (pct >= 0.55) vScore = 3;
          else vScore = 2;
        }
      }

      // Margin Score (J)
      let mScore = 0;
      if (curMargin >= 0.55) mScore = 3;
      else if (curMargin >= 0.45) mScore = 2;
      else if (curMargin >= 0.35) mScore = 1;

      // Stock Score (K)
      const webStock = safeNum(webMap.get(sku)) || 0; // Default null to 0 to prevent NaN in stock arithmetic
      let sScore = 0;
      if (fulfillment === "WEBONLY") {
        sScore = 2;
      } else {
        const dailyVelocity = (units90 || 0) / 90;
        const dos = dailyVelocity > 0 ? webStock / dailyVelocity : 999;
        if (dos <= 30) sScore = 3;
        else if (dos <= 120) sScore = 2;
        else if (dos <= 180) sScore = 1;
      }

      const totalScore = vScore + mScore + sScore;
      
      // Tiers & Logic (M, N)
      let tier = "Clearance/Archive (65% Off)";
      let vdmMarkdown = 0.65;
      let tierCode = TIER_CODES.CLEARANCE;
      if (gateCode === GATEKEEPER_CODES.GWP) { tier = "GWP Promo Hold (0% Hold)"; vdmMarkdown = 0; tierCode = TIER_CODES.GATEKEEPER; } // GWP: freeze markdown at 0%
      else if (gateCode === GATEKEEPER_CODES.NEW_LAUNCH) { tier = "New Launch (0% Hold)"; vdmMarkdown = 0; tierCode = TIER_CODES.GATEKEEPER; }
      else if (gateCode === GATEKEEPER_CODES.MAP) { tier = "3rd Party MAP Review (0% Hold)"; vdmMarkdown = 0; tierCode = TIER_CODES.GATEKEEPER; }
      else if (totalScore === 10) { tier = "Top Hero (0% Off)"; vdmMarkdown = 0; tierCode = TIER_CODES.TOP_HERO; }
      else if (totalScore >= 8) { tier = "Signature Hero (30% Off)"; vdmMarkdown = 0.30; tierCode = TIER_CODES.SIG_HERO; }
      else if (totalScore >= 6) { tier = "Proven Performer (40% Off)"; vdmMarkdown = 0.40; tierCode = TIER_CODES.PROVEN; }
      else if (totalScore >= 4) { tier = "Accelerator (50% Off)"; vdmMarkdown = 0.50; tierCode = TIER_CODES.ACCELERATOR; }
      else {
        // Score 0–3: WEBONLY strategic tier exception
        if (fulfillment === "WEBONLY") {
          if ((units90 || 0) === 0) {
            // 90-day total sales = 0 (used as the operational proxy for three consecutive 30-day zero-sales periods)
            tier = "Clearance/Archive (65% Off)";
            vdmMarkdown = 0.65;
          } else {
            // Has some sales history — hold at digital review tier
            tier = "Accelerator / Digital Review (50% Off)";
            vdmMarkdown = 0.50;
          }
        }
        // SHARED score 0–3: default Clearance/Archive (65%) already set above
      }

      const usaRow = usaMap.get(sku);
      const usaStock = usaRow ? safeNum(usaRow[uIdx[usaStockHeader]]) || 0 : 0;
      const totalStock = (safeNum(usaStock) ?? 0) + (safeNum(webStock) ?? 0);
      const shopifyQty = shopifyQtyHeader ? safeNum(row[sIdx[shopifyQtyHeader]]) || 0 : 0;
      let propPrice = compareMSRP * (1 - vdmMarkdown);
      let simNet = propPrice * (1 - affiliateRate);
      const baselineSnapshot = baselineMap.get(sku);
      const baselinePrice = baselineSnapshot && mathGuard(baselineSnapshot.price) ? baselineSnapshot.price : (price || 0);
      const baselineCompareMsrp = baselineSnapshot && mathGuard(baselineSnapshot.compareMsrp) ? baselineSnapshot.compareMsrp : compareMSRP;
      const _storedBracket = baselineSnapshot ? safeStr(baselineSnapshot.bracket) : null;
      const _canonicalBrackets = Object.values(VDM_CONFIG.BRACKET_NAMES);
      const baselineBracket = (_storedBracket && _canonicalBrackets.indexOf(_storedBracket) !== -1)
        ? _storedBracket
        : getStorefrontBracket(baselinePrice, baselineCompareMsrp);
      
      let stackMargin = 0;
      let guardrail = "✓ SAFE";
      let guardrailCode = GUARDRAIL_CODES.SAFE;
      const isCostMissing = (cost === null || cost === 0) && !isZeroCostPermitted(gateCode);

      if (isCostMissing) {
        guardrail = "❌ BLOCKED (Missing Cost)";
        guardrailCode = GUARDRAIL_CODES.ERR_MISSING_COST;
      } else if (curMargin < 0) {
        guardrail = "❌ BLOCKED (Negative Base Margin)";
        guardrailCode = GUARDRAIL_CODES.ERR_NEGATIVE_MARGIN;
      } else if (mathGuard(simNet, cost)) {
        stackMargin = simNet === 0 ? 0 : (simNet - cost) / simNet;
        if (stackMargin < VDM_CONFIG.PROFIT_FLOOR_GUARDRAIL) {
          guardrail = "❌ BLOCKED (Margin Floor Violator)";
          guardrailCode = GUARDRAIL_CODES.ERR_MARGIN_FLOOR_VIOLATOR;
        }
      }
      
      const curTierLabel = curMarkdown === 0 ? "Full MSRP" : (curMarkdown <= 0.19 ? "Promo Tier 1 (10-15%)" : (curMarkdown <= 0.35 ? "Promo Tier 2 (20-25%)" : (curMarkdown <= 0.55 ? "Promo Tier 3 (40-50%)" : "Clearance")));
      
      // X: Governance Override
      const b2b30DSales = usaRow && usaSalesHeader ? safeNum(usaRow[uIdx[usaSalesHeader]]) || 0 : 0;
      let migration = (vdmMarkdown > curMarkdown) ? "🚨 Deepen Discount" : "📈 Price Recovery/Lift";
      if (vdmMarkdown === curMarkdown) migration = "✓ Price Hold";

      // THE FIX: Intercept the text AND revert the math
      const b2bHoldActive = fulfillment === "SHARED" && (vdmMarkdown >= 0.50) && usaStock >= b2bReserveMin && b2b30DSales > 0;
      if (guardrailCode === GUARDRAIL_CODES.SAFE && b2bHoldActive) {
        migration = "⚠️ HOLD: B2B Volume Stable";
        vdmMarkdown = curMarkdown; // Revert markdown to match current live site
        tier = "B2B Protection Hold"; // Change tier name
        guardrail = "⚠️ B2B HOLD";
        guardrailCode = GUARDRAIL_CODES.WARN_B2B_HOLD;
        
        propPrice = compareMSRP * (1 - vdmMarkdown);
        simNet = propPrice * (1 - affiliateRate);
        stackMargin = simNet === 0 ? 0 : (simNet - cost) / simNet;
      }

      const targetBracket = getStorefrontBracket(propPrice, compareMSRP);
      let catalogShiftDirection = "✓ Price Hold";
      if (propPrice > baselinePrice) {
        catalogShiftDirection = "📈 Price Recovery / Lift";
      } else if (propPrice === baselinePrice) {
        catalogShiftDirection = "✓ Price Hold";
      } else if (propPrice < baselinePrice && vdmMarkdown < 0.65) {
        catalogShiftDirection = "📉 Controlled Deepen";
      } else if (propPrice < baselinePrice && vdmMarkdown >= 0.65) {
        catalogShiftDirection = "🚨 Clearance Escalation";
      }

      // Queue Assignment
      // Queue 1A: Negative base margin audit
      // Queue 1B: Simulated checkout margin guardrail (<20% => ❌ BLOCKED)
      // Queue 2: WEBONLY digital review (total score 0–3)
      // Queue 3: SHARED clearance/liquidation (total score 0–3)
      let actionQueue = "";
      let queueCode = QUEUE_CODES.NONE;
      if (guardrailCode === GUARDRAIL_CODES.ERR_MISSING_COST) {
        actionQueue = "Queue 1A: Data Error (Missing Cost)";
        queueCode = QUEUE_CODES.QUEUE_1A_COST;
        stats.missingCost++;
      } else if (guardrailCode === GUARDRAIL_CODES.ERR_NEGATIVE_MARGIN) {
        actionQueue = "Queue 1A: ❌ BLOCKED — Negative Base Margin Audit";
        queueCode = QUEUE_CODES.QUEUE_1A_MARGIN;
        stats.negativeMarginAudits++;
      } else if (guardrailCode === GUARDRAIL_CODES.ERR_MARGIN_FLOOR_VIOLATOR) {
        actionQueue = "Queue 1B: ❌ BLOCKED — Simulated Checkout Margin <20%";
        queueCode = QUEUE_CODES.QUEUE_1B_FLOOR;
        stats.blockedByMargin++;
      } else if (guardrailCode === GUARDRAIL_CODES.WARN_B2B_HOLD) {
        stats.b2bHolds++;
      } else if (fulfillment === "WEBONLY" && totalScore <= 3) {
        actionQueue = "Queue 2: WEBONLY Digital Review";
        queueCode = QUEUE_CODES.QUEUE_2_WEBONLY;
      } else if (fulfillment === "SHARED" && totalScore <= 3) {
        actionQueue = "Queue 3: SHARED Clearance / Liquidation";
        queueCode = QUEUE_CODES.QUEUE_3_CLEARANCE;
      }

      stats.total++;
      results.push([
        sku, gate, gateCode, fulfillment, cost, price, baselinePrice, compareMSRP, curMarkdown, curMargin, units90 || 0, vScore, mScore, sScore, totalScore, tier, tierCode, vdmMarkdown, baselineBracket, targetBracket, catalogShiftDirection, totalStock, webStock, shopifyQty, shopifyQty - webStock, propPrice, simNet, stackMargin, guardrail, guardrailCode, curTierLabel, migration, propPrice - (price || 0), stackMargin - curMargin, actionQueue, queueCode
      ]);
    });

    // Log Data Health Results
    Logger.log(`[SYNC COMPLETE ${VDM_CONFIG.VERSION}] Processed ${stats.total} SKUs.`);
    if (stats.missingCost > 0) Logger.log(`[WARN] ${stats.missingCost} SKUs are missing cost data (Waterfall failed).`);
    if (stats.negativeMarginAudits > 0) Logger.log(`[WARN] ${stats.negativeMarginAudits} SKUs have negative base margin.`);
    if (stats.missingInventory > 0) Logger.log(`[WARN] ${stats.missingInventory} physical inventory SKUs are unmapped to the active catalog.`);
    if (stats.blockedByMargin > 0) Logger.log(`[INFO] ${stats.blockedByMargin} SKUs blocked from target discount by profit guardrails.`);
    if (stats.b2bHolds > 0) Logger.log(`[INFO] ${stats.b2bHolds} SKUs held for B2B reserve protection.`);
    if (stats.fulfillmentFallbackCount > 0) Logger.log(`[WARN] ${stats.fulfillmentFallbackCount} SKUs used fulfillment type fallback.`);

    // 2. Batch Write
    dashSheet.clear().clearFormats();
    const dashboardHeaders = ["SKU Anchor Key", "Gatekeeper Status", "Gatekeeper Code", "Fulfillment Tag", "Resolved Cost Base", "Live Storefront Price", "Baseline Storefront Price", "Live Compare MSRP", "Active Storefront Markdown Depth %", "Current Gross Margin %", "Raw 90D Retail Velocity", "Retail Velocity Score Component", "Margin Score Component", "Retail Stock Score Component", "Total Composite Score", "Target Strategic Tier", "Tier Code", "VDM Markdown Depth %", "Baseline Storefront Bracket", "Target VDM Bracket", "Catalog Shift Direction", "Total On-Hand Warehouse Stock", "EEI Web Warehouse On Hand Stock", "Live Storefront Shopify Qty", "Asynchronous Inventory Drift Tracker", "New Proposed Storefront Price", "Simulated Checkout Net Price", "Final Simulated Stacked Margin %", "Profit Guardrail Status Alert", "Guardrail Code", "Current Equivalent Storefront Tier", "Pricing Migration Status", "Retail Price Shift ($)", "Net Margin Change %", "Action Queue", "Queue Code"];
    
    const headerWidth = dashboardHeaders.length;
    const guardrailColumnIndex = dashboardHeaders.indexOf("Profit Guardrail Status Alert") + 1;
    const rowCount = (results && results.length) ? results.length : 0;
    const headerRange = dashSheet.getRange(1, 1, 1, headerWidth);
    headerRange.setValues([dashboardHeaders]);
    applyHeaderStyle(headerRange);
    if (rowCount > 0) {
      dashSheet.getRange(2, 1, rowCount, headerWidth).setValues(results);
      dashSheet.getRange(2, 1, rowCount, 1).setNumberFormat("@");
      applyConditionalFormatting(dashSheet, rowCount, headerWidth, guardrailColumnIndex);

      // Automated Recovery Point Sync
      const backupSheet = getOrCreateSheet(VDM_CONFIG.TABS.BACKUP, true);
      backupSheet.clear();
      dashSheet.getDataRange().copyTo(backupSheet.getRange(1,1));
    }
    dashSheet.setFrozenRows(1);

    return { rows: results, headers: dashboardHeaders, stats };
  } catch (e) {
    logError("MatrixEngine", e);
    throw e;
  }
}

function applyConditionalFormatting(sheet, rowCount, colCount, guardrailColumnIndex) {
  const range = sheet.getRange(2, 1, rowCount, colCount);
  sheet.clearConditionalFormatRules();
  const guardrailColumnLetter = columnToLetter(guardrailColumnIndex);
  
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=REGEXMATCH($${guardrailColumnLetter}2,"^❌ BLOCKED")`)
      .setBackground(VDM_CONFIG.DESIGN.ALERT_BREACH_BG)
      .setFontColor(VDM_CONFIG.DESIGN.ALERT_BREACH_TEXT)
      .setBold(true)
      .setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$B2="⚠️ Active GWP Promo"')
      .setBackground(VDM_CONFIG.DESIGN.ALERT_GWP_BG)
      .setFontColor(VDM_CONFIG.DESIGN.ALERT_GWP_TEXT)
      .setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$B2="New Launch"')
      .setBackground(VDM_CONFIG.DESIGN.ALERT_LAUNCH_BG)
      .setFontColor(VDM_CONFIG.DESIGN.ALERT_LAUNCH_TEXT)
      .setRanges([range]).build()
  ];
  sheet.setConditionalFormatRules(rules);
}

function columnToLetter(column) {
  let index = column;
  let letter = "";
  while (index > 0) {
    const remainder = (index - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    index = Math.floor((index - 1) / 26);
  }
  return letter;
}