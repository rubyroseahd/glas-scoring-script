/**
 * MODULE 3: MATRIX ENGINE
 */

function executeDashboardRefresh() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashSheet = ss.getSheetByName(VDM_CONFIG.TABS.DASHBOARD);
    
    // 1. Memory Load: Load all raw data into lookup objects
    const shopifyData = ss.getSheetByName(VDM_CONFIG.TABS.RAW_SHOPIFY).getDataRange().getValues();
    const salesData = ss.getSheetByName(VDM_CONFIG.TABS.RAW_SALES).getDataRange().getValues();
    const usaData = ss.getSheetByName(VDM_CONFIG.TABS.RAW_EEI_USA).getDataRange().getValues();
    const webData = ss.getSheetByName(VDM_CONFIG.TABS.RAW_EEI_WEB).getDataRange().getValues();
    const costData = ss.getSheetByName(VDM_CONFIG.TABS.MASTER_COST).getDataRange().getValues();
    const settingsData = ss.getSheetByName(VDM_CONFIG.TABS.SETTINGS).getDataRange().getValues();

    const sIdx = getHeaderMap(shopifyData[0]);
    const vIdx = getHeaderMap(salesData[0]);
    const uIdx = getHeaderMap(usaData[0]);
    const wIdx = getHeaderMap(webData[0]);
    const cIdx = getHeaderMap(costData[0]);

    const salesMap = new Map(salesData.slice(1).map(r => [safeStr(r[0]), safeNum(r[vIdx["Net items sold"]])]));
    const usaMap = new Map(usaData.slice(1).map(r => [safeStr(r[0]), r]));
    const webMap = new Map(webData.slice(1).map(r => [safeStr(r[0]), safeNum(r[wIdx["EEI Web Warehouse On Hand Stock"]])]));
    const costMap = new Map(costData.slice(1).map(r => [safeStr(r[0]), safeNum(r[cIdx["Resolved Cost"]])]));
    
    // Load Registries
    const gwpSet = new Set(settingsData.map(r => safeStr(r[0]).toUpperCase()));
    const launchSet = new Set(settingsData.map(r => safeStr(r[1]).toUpperCase()));
    const mapBrands = settingsData.slice(1).map(r => safeStr(r[2]).toUpperCase()).filter(v => v);
    const affiliateRate = settingsData.length > 1 ? safeNum(settingsData[1][4]) : 0.15; // Fallback to 15%

    // Velocity Percentile Setup
    const salesArray = Array.from(salesMap.values()).map(v => safeNum(v)).filter(v => v > 1).sort((a,b) => a-b);

    const results = [];
    shopifyData.slice(1).forEach(row => {
      const sku = row[0];
      const vendor = safeStr(row[sIdx["Vendor"]]).toUpperCase();
      
      // A: SKU Anchor
      // B: Gatekeeper
      let gate = "None";
      if (gwpSet.has(sku)) gate = "⚠️ Active GWP Promo";
      else if (launchSet.has(sku)) gate = "New Launch";
      else if (mapBrands.some(b => vendor.includes(b))) gate = "3rd Party MAP";

      const fulfillment = safeStr(row[sIdx["Fulfillment service"]]) || "SHARED";
      const cost = safeNum(costMap.get(sku));
      const price = safeNum(row[sIdx["Variant Price"]]);
      const rawCompare = safeNum(row[sIdx["Variant Compare At Price"]]);
      const compareMSRP = (rawCompare === 0 || isNaN(rawCompare)) ? price : rawCompare;
      const curMarkdown = compareMSRP === price ? 0 : (compareMSRP - price) / compareMSRP;
      const curMargin = price === 0 ? 0 : (price - cost) / price;
      
      // Velocity Score (I)
      const units90 = safeNum(salesMap.get(sku));
      let vScore = 0;
      if (units90 === 1) {
        vScore = 1;
      } else if (units90 > 1) {
        const den = salesArray.length - 1;
        const rank = den > 0 ? salesArray.filter(v => v < units90).length / den : 0;
        if (rank >= 0.80) vScore = 4;
        else if (rank >= 0.55) vScore = 3;
        else vScore = 2;
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
        const dailyVelocity = safeNum(units90 / 90);
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
      const usaStock = usaRow ? safeNum(usaRow[uIdx["EEI USA Warehouse On Hand Stock"]]) : 0;
      const totalStock = usaStock + webStock;
      const shopifyQty = safeNum(row[sIdx["Variant Inventory Qty"]]);
      const propPrice = compareMSRP * (1 - vdmMarkdown);
      const simNet = propPrice * (1 - affiliateRate);
      const stackMargin = simNet === 0 ? 0 : (simNet - cost) / simNet;
      const guardrail = stackMargin < 0.20 ? "❌ BLOCKED" : "✓ SAFE";
      
      const curTierLabel = curMarkdown === 0 ? "Full MSRP" : (curMarkdown <= 0.19 ? "Promo Tier 1 (10-15%)" : (curMarkdown <= 0.35 ? "Promo Tier 2 (20-25%)" : (curMarkdown <= 0.55 ? "Promo Tier 3 (40-50%)" : "Clearance")));
      
      // X: Governance Override
      const b2b30DSales = usaRow ? safeNum(usaRow[uIdx["Sales Past 30 Days"]]) : 0;
      let migration = (vdmMarkdown > curMarkdown) ? "🚨 Deepen Discount" : "📈 Price Recovery/Lift";
      if (vdmMarkdown === curMarkdown) migration = "✓ Price Hold";
      if (fulfillment === "SHARED" && (vdmMarkdown >= 0.50) && usaStock >= 500 && b2b30DSales > 0) {
        migration = "⚠️ HOLD: B2B Volume Stable";
      }

      results.push([
        sku, gate, fulfillment, cost, price, compareMSRP, curMarkdown, curMargin, vScore, mScore, sScore, totalScore, tier, vdmMarkdown, totalStock, webStock, shopifyQty, shopifyQty - webStock, propPrice, simNet, stackMargin, guardrail, curTierLabel, migration, propPrice - price, stackMargin - curMargin
      ]);
    });

    // 2. Batch Write
    dashSheet.clear().clearFormats();
    const headers = [
      "SKU Anchor Key", "Gatekeeper Status", "Fulfillment Tag", "Resolved Cost Base", "Live Storefront Price",
      "Live Compare MSRP", "Active Storefront Markdown Depth %", "Current Gross Margin %", "Retail Velocity Score Component",
      "Margin Score Component", "Retail Stock Score Component", "Total Composite Score", "Target Strategic Tier",
      "VDM Markdown Depth %", "Total On-Hand Warehouse Stock", "EEI Web Warehouse On Hand Stock", 
      "Live Storefront Shopify Qty", "Asynchronous Inventory Drift Tracker", "New Proposed Storefront Price",
      "Simulated Checkout Net Price", "Final Simulated Stacked Margin %", "Profit Guardrail Status Alert",
      "Current Equivalent Storefront Tier", "Pricing Migration Status", "Retail Price Shift ($)", "Net Margin Change %"
    ];
    
    const headerRange = dashSheet.getRange(1, 1, 1, 26);
    headerRange.setValues([headers]);
    applyHeaderStyle(headerRange);
    if (results.length > 0) {
      dashSheet.getRange(2, 1, results.length, 26).setValues(results);
      dashSheet.getRange(2, 1, results.length, 1).setNumberFormat("@");
    }
    
    applyConditionalFormatting(dashSheet, results.length);
    dashSheet.setFrozenRows(1);
  } catch (e) {
    logError("MatrixEngine", e);
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