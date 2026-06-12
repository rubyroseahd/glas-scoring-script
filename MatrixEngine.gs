/**
 * MODULE 3: MATRIX ENGINE
 */

function executeDashboardRefresh() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dash = ss.getSheetByName(VDM_CONFIG.TABS.DASHBOARD);
    const shopify = ss.getSheetByName(VDM_CONFIG.TABS.RAW_SHOPIFY);
    const shopifyData = shopify.getDataRange().getValues();
    
    if (shopifyData.length < 2) return;
    const skus = shopifyData.slice(1).map(r => [r[0]]);
    
    dash.clear().clearFormats();
    const headers = [
      "SKU Anchor Key", "Gatekeeper Status", "Fulfillment Tag", "Resolved Cost Base", "Live Storefront Price",
      "Live Compare MSRP", "Active Storefront Markdown Depth %", "Current Gross Margin %", "Retail Velocity Score Component",
      "Margin Score Component", "Retail Stock Score Component", "Total Composite Score", "Target Strategic Tier",
      "VDM Markdown Depth %", "Total On-Hand Warehouse Stock", "EEI Web Warehouse On Hand Stock", 
      "Live Storefront Shopify Qty", "Asynchronous Inventory Drift Tracker", "New Proposed Storefront Price",
      "Simulated Checkout Net Price", "Final Simulated Stacked Margin %", "Profit Guardrail Status Alert",
      "Current Equivalent Storefront Tier", "Pricing Migration Status", "Retail Price Shift ($)", "Net Margin Change %"
    ];
    
    const headerRange = dash.getRange(1, 1, 1, 26);
    headerRange.setValues([headers]);
    applyHeaderStyle(headerRange);
    dash.getRange(2, 1, skus.length, 1).setValues(skus).setNumberFormat("@");

    const ctrl = `'${VDM_CONFIG.TABS.SETTINGS}'`;
    const rShop = VDM_CONFIG.TABS.RAW_SHOPIFY;
    const rSales = VDM_CONFIG.TABS.RAW_SALES;
    const rUsa = VDM_CONFIG.TABS.RAW_EEI_USA;
    const rWeb = VDM_CONFIG.TABS.RAW_EEI_WEB;
    const rCost = VDM_CONFIG.TABS.MASTER_COST;

    const formulas = skus.map((_, i) => {
      const r = i + 2;
      return [
        `=IFS(ISNUMBER(MATCH(A${r}, ${ctrl}!$A:$A, 0)), "⚠️ Active GWP Promo", ISNUMBER(MATCH(A${r}, ${ctrl}!$B:$B, 0)), "New Launch", SUMPRODUCT(--ISNUMBER(SEARCH(${ctrl}!$C$2:$C$50, VLOOKUP(A${r}, ${rShop}!$A:$G, 7, 0)))), "3rd Party MAP", TRUE, "None")`,
        `=IFERROR(VLOOKUP(A${r}, ${rShop}!$A:$G, 6, 0), "SHARED")`,
        `=VLOOKUP(A${r}, ${rCost}!$A:$B, 2, 0)`,
        `=VLOOKUP(A${r}, ${rShop}!$A:$G, 4, 0)`,
        `=IF(OR(ISBLANK(VLOOKUP(A${r}, ${rShop}!$A:$E, 5, 0)), VLOOKUP(A${r}, ${rShop}!$A:$E, 5, 0)=0), E${r}, VLOOKUP(A${r}, ${rShop}!$A:$E, 5, 0))`,
        `=IF(F${r}=E${r}, 0, (F${r} - E${r}) / F${r})`,
        `=IFERROR((E${r} - D${r}) / E${r}, 0)`,
        `=IF(IFERROR(VLOOKUP(A${r}, ${rSales}!$A:$Z, 4, 0), 0)=0, 0, IF(VLOOKUP(A${r}, ${rSales}!$A:$Z, 4, 0)=1, 1, IF(PERCENTRANK.INC(FILTER(${rSales}!$D$2:$D, ${rSales}!$D$2:$D>1), VLOOKUP(A${r}, ${rSales}!$A:$Z, 4, 0))>=0.80, 4, IF(PERCENTRANK.INC(FILTER(${rSales}!$D$2:$D, ${rSales}!$D$2:$D>1), VLOOKUP(A${r}, ${rSales}!$A:$Z, 4, 0))>=0.55, 3, 2))))`,
        `=IFERROR(IFS(H${r}>=0.55, 3, H${r}>=0.45, 2, H${r}>=0.35, 1, TRUE, 0), 0)`,
        `=IF(C${r}="WEBONLY", 2, LET(sls, IFERROR(VLOOKUP(A${r}, ${rSales}!$A:$Z, 4, 0), 0), dos, IF(sls=0, 999, P${r}/(sls/90)), IFS(dos<=30, 3, dos<=120, 2, dos<=180, 1, TRUE, 0)))`,
        `=SUM(I${r}, J${r}, K${r})`,
        `=IFS(B${r}="New Launch", "New Launch (0% Hold)", B${r}="3rd Party MAP", "3rd Party MAP Review (0% Hold)", L${r}=10, "Top Hero (0% Off)", L${r}>=8, "Signature Hero (30% Off)", L${r}>=6, "Proven Performer (40% Off)", L${r}>=4, "Accelerator (50% Off)", TRUE, "Clearance/Archive (65% Off)")`,
        `=IF(M${r}="Signature Hero (30% Off)", 0.30, IF(M${r}="Proven Performer (40% Off)", 0.40, IF(M${r}="Accelerator (50% Off)", 0.50, IF(M${r}="Clearance/Archive (65% Off)", 0.65, 0))))`,
        `=IFERROR(VLOOKUP(A${r}, ${rUsa}!$A:$M, 3, 0), 0) + IFERROR(VLOOKUP(A${r}, ${rWeb}!$A:$M, 3, 0), 0)`,
        `=IFERROR(VLOOKUP(A${r}, ${rWeb}!$A:$M, 3, 0), 0)`,
        `=IFERROR(VLOOKUP(A${r}, ${rShop}!$A:$Z, 8, 0), 0)`,
        `=Q${r} - P${r}`,
        `=F${r} * (1 - N${r})`,
        `=S${r} * (1 - ${ctrl}!$E$2)`,
        `=IFERROR((T${r} - D${r}) / T${r}, 0)`,
        `=IF(U${r}<0.20, "❌ BLOCKED", "✓ SAFE")`,
        `=IFS(G${r}=0, "Full MSRP", G${r}<=0.19, "Promo Tier 1 (10-15%)", G${r}<=0.35, "Promo Tier 2 (20-25%)", G${r}<=0.55, "Promo Tier 3 (40-50%)", TRUE, "Clearance")`,
        `=IFS(W${r}=LEFT(M${r}, LEN(W${r})), "✓ Price Hold", AND(C${r}="SHARED", OR(LEFT(M${r},9)="Clearance", LEFT(M${r},11)="Accelerator"), IFERROR(VLOOKUP(A${r}, ${rUsa}!$A:$M, 3, 0), 0)>=500, IFERROR(VLOOKUP(A${r}, ${rUsa}!$A:$O, 4, 0), 0)>0), "⚠️ HOLD: B2B Volume Stable", N${r}>G${r}, "🚨 Deepen Discount", TRUE, "📈 Price Recovery/Lift")`,
        `=S${r} - E${r}`,
        `=U${r} - H${r}`
      ];
    });

    dash.getRange(2, 2, formulas.length, 25).setFormulas(formulas);
    applyConditionalFormatting(dash, formulas.length);
    dash.setFrozenRows(1);
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