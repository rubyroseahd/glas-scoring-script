/**
 * MODULE 3: MATRIX ENGINE
 */

function executeDashboardRefresh() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashSheet = ss.getSheetByName(VDM_CONFIG.TABS.DASHBOARD);
    
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

    const salesMap = new Map(salesData.slice(1).map(r => [safeStr(r[0]), safeNum(r[vIdx["NET ITEMS SOLD"]])]));
    const usaMap = new Map(usaData.slice(1).map(r => [safeStr(r[0]), r]));
    const webMap = new Map(webData.slice(1).map(r => [safeStr(r[0]), safeNum(r[wIdx["EEI WEB WAREHOUSE ON HAND STOCK"]])]));
    const costMap = new Map(costData.slice(1).map(r => [safeStr(r[0]), safeNum(r[cIdx["RESOLVED COST"]])]));
    
    // Load Registries
    const gwpSet = new Set(settingsData.slice(1).map(r => safeStr(r[0]).toUpperCase())); // Skip header row
    const launchSet = new Set(settingsData.slice(1).map(r => safeStr(r[1]).toUpperCase())); // Skip header row
    const mapBrands = settingsData.slice(1).map(r => safeStr(r[2]).toUpperCase()).filter(v => v); // Skip header row
    const affiliateRate = (settingsData.length > 1 && safeNum(settingsData[1][4]) !== null) ? safeNum(settingsData[1][4]) : 0.15; // Fallback to 15%

    // Velocity Percentile Setup
    const salesArray = Array.from(salesMap.values()).filter(v => v !== null && v > 1).sort((a,b) => a-b);
      
    // Data Health Tracking
    const stats = {
      total: 0,
      missingCost: 0,
      missingInventory: 0,
      blockedByMargin: 0
    };

    const results = [];
    shopifyData.slice(1).forEach(row => {
      const sku = row[0];
      const vendor = safeStr(row[sIdx["VENDOR"]]).toUpperCase();
      
      // A: SKU Anchor
      // B: Gatekeeper
      let gate = "None";
      if (gwpSet.has(sku)) gate = "⚠️ Active GWP Promo";
      else if (launchSet.has(sku)) gate = "New Launch";
      else if (mapBrands.some(b => vendor.includes(b))) gate = "3rd Party MAP";

      const fulfillment = usaMap.has(sku) ? "SHARED" : "WEBONLY";
      const cost = safeNum(costMap.get(sku));
      const price = safeNum(row[sIdx["VARIANT PRICE"]]);
      if (cost === null) stats.missingCost++;
      const rawCompare = safeNum(row[sIdx["VARIANT COMPARE AT PRICE"]]);
      
      const compareMSRP = (rawCompare === 0 || rawCompare === null) ? (price || 0) : rawCompare;
      const curMarkdown = (compareMSRP === price || compareMSRP === 0) ? 0 : (compareMSRP - (price || 0)) / compareMSRP;
      const curMargin = (price === 0 || price === null || cost === null) ? 0 : (price - cost) / price;
      
      // Velocity Score (I) - Ensure sIdx["Net items sold"] is valid
      const units90 = safeNum(salesMap.get(sku));
      let vScore = 0;
      if (units90 !== null) {
        if (units90 === 1) {
          vScore = 1;
        } else if (units90 > 1) {
          const den = salesArray.length - 1;
          const rank = den > 0 ? salesArray.filter(v => v < units90).length / den : 0;
          if (rank >= 0.80) vScore = 4;
          else if (rank >= 0.55) vScore = 3;
          else vScore = 2;
        }
      }

      // Margin Score (J)
      let mScore = 0;
      if (curMargin >= 0.55) mScore = 3;
      else if (curMargin >= 0.45) mScore = 2;
      else if (curMargin >= 0.35) mScore = 1;

      // Stock Score (K)
      const webStock = safeNum(webMap.get(sku));
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
      if (gate === "New Launch") { tier = "New Launch (0% Hold)"; vdmMarkdown = 0; }
      else if (gate === "3rd Party MAP") { tier = "3rd Party MAP Review (0% Hold)"; vdmMarkdown = 0; }
      else if (totalScore === 10) { tier = "Top Hero (0% Off)"; vdmMarkdown = 0; }
      else if (totalScore >= 8) { tier = "Signature Hero (30% Off)"; vdmMarkdown = 0.30; }
      else if (totalScore >= 6) { tier = "Proven Performer (40% Off)"; vdmMarkdown = 0.40; }
      else if (totalScore >= 4) { tier = "Accelerator (50% Off)"; vdmMarkdown = 0.50; }

      const usaRow = usaMap.get(sku);
      const usaStock = usaRow ? safeNum(usaRow[uIdx["EEI USA WAREHOUSE ON HAND STOCK"]]) || 0 : 0;
      const totalStock = usaStock + webStock;
      if (!usaRow && fulfillment === "SHARED") stats.missingInventory++;
      const shopifyQty = safeNum(row[sIdx["VARIANT INVENTORY QTY"]]) || 0;
      let propPrice = compareMSRP * (1 - vdmMarkdown);
      let simNet = propPrice * (1 - affiliateRate);
      
      let stackMargin = 0;
      let guardrail = "✓ SAFE";

      if (cost === null) {
        guardrail = "DATA_ERROR";
      } else if (mathGuard(simNet, cost)) {
        stackMargin = simNet === 0 ? 0 : (simNet - cost) / simNet;
        if (stackMargin < 0.20) guardrail = "❌ BLOCKED";
        if (guardrail === "❌ BLOCKED") stats.blockedByMargin++;
      }
      
      const curTierLabel = curMarkdown === 0 ? "Full MSRP" : (curMarkdown <= 0.19 ? "Promo Tier 1 (10-15%)" : (curMarkdown <= 0.35 ? "Promo Tier 2 (20-25%)" : (curMarkdown <= 0.55 ? "Promo Tier 3 (40-50%)" : "Clearance")));
      
      // X: Governance Override
      const b2b30DSales = usaRow ? safeNum(usaRow[uIdx["SALES PAST 30 DAYS"]]) || 0 : 0;
      let migration = (vdmMarkdown > curMarkdown) ? "🚨 Deepen Discount" : "📈 Price Recovery/Lift";
      if (vdmMarkdown === curMarkdown) migration = "✓ Price Hold";

      // THE FIX: Intercept the text AND revert the math
      if (fulfillment === "SHARED" && (vdmMarkdown >= 0.50) && usaStock >= 500 && b2b30DSales > 0) {
        migration = "⚠️ HOLD: B2B Volume Stable";
        vdmMarkdown = curMarkdown; // Revert markdown to match current live site
        tier = "B2B Protection Hold"; // Change tier name
        
        // Recalculate safe pricing
        propPrice = compareMSRP * (1 - vdmMarkdown);
        simNet = propPrice * (1 - affiliateRate);
        stackMargin = simNet === 0 ? 0 : (simNet - cost) / simNet;
        guardrail = stackMargin < 0.20 ? "❌ BLOCKED" : "✓ SAFE";
      }

      stats.total++;
      results.push([
        sku, gate, fulfillment, cost, price, compareMSRP, curMarkdown, curMargin, vScore, mScore, sScore, totalScore, tier, vdmMarkdown, totalStock, webStock, shopifyQty, shopifyQty - webStock, propPrice, simNet, stackMargin, guardrail, curTierLabel, migration, propPrice - (price || 0), stackMargin - curMargin
      ]);
    });

    // Log Data Health Results
    Logger.log(`[SYNC COMPLETE] Processed ${stats.total} SKUs.`);
    if (stats.missingCost > 0) Logger.log(`[WARN] ${stats.missingCost} SKUs are missing cost data (Waterfall failed).`);
    if (stats.missingInventory > 0) Logger.log(`[WARN] ${stats.missingInventory} SHARED SKUs missing from USA Warehouse file.`);
    if (stats.blockedByMargin > 0) Logger.log(`[INFO] ${stats.blockedByMargin} SKUs blocked from target discount by profit guardrails.`);

    // 2. Batch Write
    dashSheet.clear().clearFormats();
    const dashboardHeaders = [
      "SKU Anchor Key", "Gatekeeper Status", "Fulfillment Tag", "Resolved Cost Base", "Live Storefront Price",
      "Live Compare MSRP", "Active Storefront Markdown Depth %", "Current Gross Margin %", "Retail Velocity Score Component",
      "Margin Score Component", "Retail Stock Score Component", "Total Composite Score", "Target Strategic Tier",
      "VDM Markdown Depth %", "Total On-Hand Warehouse Stock", "EEI Web Warehouse On Hand Stock", 
      "Live Storefront Shopify Qty", "Asynchronous Inventory Drift Tracker", "New Proposed Storefront Price",
      "Simulated Checkout Net Price", "Final Simulated Stacked Margin %", "Profit Guardrail Status Alert",
      "Current Equivalent Storefront Tier", "Pricing Migration Status", "Retail Price Shift ($)", "Net Margin Change %"
    ];
    
    const headerRange = dashSheet.getRange(1, 1, 1, 26);
    headerRange.setValues([dashboardHeaders]);
    applyHeaderStyle(headerRange);
    if (results.length > 0 && results[0].length > 0) { // Ensure results array is not empty and has columns
      dashSheet.getRange(2, 1, results.length, 26).setValues(results);
      dashSheet.getRange(2, 1, results.length, 1).setNumberFormat("@");
    }
    
    applyConditionalFormatting(dashSheet, results.length);
    dashSheet.setFrozenRows(1);

    return { rows: results, headers: dashboardHeaders };
  } catch (e) {
    logError("MatrixEngine", e);
    throw e;
  }
}

function applyConditionalFormatting(sheet, rowCount) {
  const range = sheet.getRange(2, 1, rowCount, 26);
  sheet.clearConditionalFormatRules();
  
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$V2="❌ BLOCKED"')
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