/**
 * MODULE 5: UI INTERFACE
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("⚡ VDM Engine")
    .addItem("1. Run Full System Pricing & Sync", "executeDashboardRefresh")
    .addSeparator()
    .addItem("2. Refresh Action Items Hub Only", "workflowRefreshActionHubOnly")
    .addItem("3. Refresh Tier Summary Only", "workflowRefreshTierSummaryOnly")
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu("Advanced Diagnostics")
        .addItem("Freeze Pre-Campaign Baseline Snapshot", "workflowFreezeBaselineSnapshot")
    )
    .addToUi();
}

function workflowFreezeBaselineSnapshot() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const source = ss.getSheetByName(VDM_CONFIG.TABS.SHOPIFY_EXPORT);
    if (!source || source.getLastRow() < 2) {
      throw new Error("shopify_export does not contain active catalog rows.");
    }

    const data = source.getDataRange().getValues();
    const idx = getHeaderMap(data[0]);
    const skuHeader = getFirstAvailableHeader(idx, ["SKU_ANCHOR", "VARIANT SKU", "SKU"]);
    const statusHeader = getFirstAvailableHeader(idx, ["STATUS"]);
    const priceHeader = getFirstAvailableHeader(idx, ["VARIANT PRICE", "PRICE"]);
    const compareHeader = findFirstAvailableHeader(idx, ["VARIANT COMPARE AT PRICE", "COMPARE AT PRICE"]);

    const rows = [];
    data.slice(1).forEach(row => {
      const status = safeStr(row[idx[statusHeader]]).toLowerCase();
      if (status !== "active") return;
      const sku = safeStr(row[idx[skuHeader]]).toUpperCase();
      if (!sku) return;
      const livePrice = safeNum(row[idx[priceHeader]]) || 0;
      const compareRaw = compareHeader ? safeNum(row[idx[compareHeader]]) : null;
      const resolvedMsrp = compareRaw !== null && compareRaw > livePrice ? compareRaw : livePrice;
      const markdown = resolvedMsrp > 0 ? (resolvedMsrp - livePrice) / resolvedMsrp : 0;
      rows.push([sku, livePrice, resolvedMsrp, markdown, new Date()]);
    });

    const baseline = getOrCreateSheet(VDM_CONFIG.TABS.BASELINE);
    const payload = [["SKU Anchor", "Live Price", "Compare MSRP", "Active Markdown %", "Captured Timestamp"]].concat(rows);
    baseline.clearContents();
    baseline.getRange(1, 1, payload.length, payload[0].length).setValues(payload);
    ui.alert("Baseline snapshot frozen.");
  } catch (e) {
    ui.alert("Baseline snapshot failed: " + e.message);
  }
}
