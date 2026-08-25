/**
 * MODULE 5: UI INTERFACE
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('EEI Pricing Engine Launcher')
    .addItem('1. Full System Sync (Standard)', 'executeFlexibleRefreshProcess')
    .addSeparator()
    .addSubMenu(ui.createMenu('2. Targeted Data Ingestion')
      .addItem('Sync Inventory Snapshot Only', 'workflowIngestInventoryOnly')
      .addItem('Refresh Commercial Metadata', 'workflowIngestMetadataOnly'))
    .addSubMenu(ui.createMenu('3. Re-calculate & Simulate')
      .addItem('Recalculate Matrix (Memory Only)', 'workflowComputeOnly'))
    .addSubMenu(ui.createMenu('4. Generate Specific Reports')
      .addItem('Update Executive Summary [03] Only', 'workflowReportSummaryOnly')
      .addItem('Generate Sync Audit [07] Only', 'workflowReportSyncOnly')
      .addItem('Refresh Master Ledger [09] Only', 'workflowReportLedgerOnly'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Advanced Diagnostics')
      .addItem('Run Pre-Flight Sanity Check', 'runPreFlightSanityCheck')
      .addItem('Freeze Pre-Campaign Baseline', 'freezePreCampaignBaseline')
      .addItem('Clear Baseline Snapshot', 'clearPreCampaignBaseline')
      .addItem('Refresh Shopify Sync Audit Only (Bypass Matrix)', 'refreshShopifySyncAuditOnly')
      .addItem('Emergency Matrix Rollback', 'rollbackToRecoveryPoint'))
    .addSeparator()
    .addItem('5. Reset Grid Architecture Logs', 'triggerNuclearArchitectureWipe')
    .addToUi();
}

function triggerNuclearArchitectureWipe() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'CRITICAL RESET REQUIRED. This will wipe all dashboards and historical logs to rebuild the system architecture. Confirm execution?', ui.ButtonSet.YES_NO);

  if (response === ui.Button.YES) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Object.values(VDM_CONFIG.TABS).forEach(tabName => {
      let sheet = ss.getSheetByName(tabName);
      if (sheet) {
        sheet.clear().clearFormats();
      } else {
        sheet = ss.insertSheet(tabName);
      }
      if (tabName.startsWith('_')) sheet.hideSheet();
    });
    
    // Initialize Settings Tab
    const settings = ss.getSheetByName(VDM_CONFIG.TABS.SETTINGS);
    const settingsHeaders = [["Active GWP SKUs", "New Launch Overrides", "MAP Restricted Brands", "B2B Reserve Min Qty", "Affiliate Coupon Rate", "Virtual SKU Prefixes"]];
    settings.getRange(1, 1, 1, settingsHeaders[0].length).setValues(settingsHeaders);
    applyHeaderStyle(settings.getRange(1, 1, 1, settingsHeaders[0].length));
    settings.getRange("E2").setValue(0.15).setNumberFormat("0.00%");
    
    // Delete specific legacy tabs that are no longer in VDM_CONFIG.TABS
    deleteSpecificLegacyTabs(ss);
    
    ui.alert("System Architecture Wiped and Rebuilt.");
  }
}

/**
 * Deletes specific legacy tabs that are no longer part of the VDM_CONFIG.TABS registry.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
 */
function deleteSpecificLegacyTabs(ss) {
  const legacyTabNames = ["[05] Warehouse Aging", "[06] MAP Compliance"];
  legacyTabNames.forEach(tabName => {
    const sheet = ss.getSheetByName(tabName);
    if (sheet) {
      ss.deleteSheet(sheet);
      Logger.log(`Deleted legacy sheet: ${tabName}`);
    }
  });
}

function runPreFlightSanityCheck() {
  const ui = SpreadsheetApp.getUi();
  try {
    const folder = DriveApp.getFolderById(getOperationalFolderId());
    validateHeaders(folder);
    ui.alert("SUCCESS: All source file headers and directory structures validated.");
  } catch (e) {
    ui.alert("SANITY CHECK FAILED: " + e.message);
  }
}

function getSyncAuditRefreshSuccessMessage() {
  return "Shopify Sync Audit refreshed for review only. No Shopify/storefront prices were changed.";
}

/**
 * Report-only diagnostic workflow: restores the dashboard state and regenerates
 * the [07] sync audit without publishing any updates to Shopify/storefront pricing.
 */
function refreshShopifySyncAuditOnly() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const state = recoverDashboardState();
    const shopifyMap = getShopifyMap();
    generateSyncAudit(ss, state.rows, getHeaderMap(state.headers), shopifyMap);
    ui.alert(getSyncAuditRefreshSuccessMessage());
  } catch (e) {
    ui.alert("Sync Refresh Failed: " + e.message);
  }
}

function rollbackToRecoveryPoint() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert("PANIC ROLLBACK: Overwrite active Matrix with last stable backup?", ui.ButtonSet.YES_NO);
  if (confirm === ui.Button.YES) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dash = getOrCreateSheet(VDM_CONFIG.TABS.DASHBOARD);
    const backup = ss.getSheetByName(VDM_CONFIG.TABS.BACKUP);
    if (!backup) throw new Error("No recovery point found.");
    dash.clear();
    backup.getDataRange().copyTo(dash.getRange(1,1));
    ui.alert("System Rollback Complete.");
  }
}

function freezePreCampaignBaseline() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(VDM_CONFIG.TABS.RAW_SHOPIFY);
    if (!sourceSheet || sourceSheet.getLastRow() < 2) throw new Error("Shopify export data not found.");

    const data = sourceSheet.getDataRange().getValues();
    const idx = getHeaderMap(data[0]);
    const skuHeader = getFirstAvailableHeader(idx, ["SKU_ANCHOR", "VARIANT SKU", "SKU"]);
    const priceHeader = getFirstAvailableHeader(idx, ["VARIANT PRICE", "PRICE"]);
    const compareHeader = findFirstAvailableHeader(idx, ["VARIANT COMPARE AT PRICE", "COMPARE AT PRICE"]);
    const capturedAt = new Date();

    const baselineRows = data.slice(1).map(row => {
      const sku = safeStr(row[idx[skuHeader]]).toUpperCase();
      if (!sku) return null;
      const livePrice = safeNum(row[idx[priceHeader]]) || 0;
      const rawCompare = compareHeader ? safeNum(row[idx[compareHeader]]) : null;
      const compareMsrp = (rawCompare === null || rawCompare === 0) ? livePrice : rawCompare;
      const markdown = compareMsrp === 0 || compareMsrp === livePrice ? 0 : (compareMsrp - livePrice) / compareMsrp;
      return [sku, livePrice, compareMsrp, markdown, getStorefrontBracket(livePrice, compareMsrp), capturedAt];
    }).filter(row => row !== null);

    const baselineSheet = getOrCreateSheet(VDM_CONFIG.TABS.BASELINE);
    baselineSheet.clear().clearFormats();
    const headers = [["SKU Anchor", "Live Price", "Compare MSRP", "Active Markdown %", "Baseline Bracket", "Captured Timestamp"]];
    baselineSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    applyHeaderStyle(baselineSheet.getRange(1, 1, 1, headers[0].length));
    if (baselineRows.length > 0) {
      baselineSheet.getRange(2, 1, baselineRows.length, headers[0].length).setValues(baselineRows);
      baselineSheet.getRange(2, 2, baselineRows.length, 2).setNumberFormat("0.00");
      baselineSheet.getRange(2, 4, baselineRows.length, 1).setNumberFormat("0.00%");
      baselineSheet.getRange(2, 6, baselineRows.length, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
    }
    baselineSheet.setFrozenRows(1);

    ui.alert("Pre-Campaign Baseline Frozen Successfully.");
  } catch (e) {
    ui.alert("Baseline Freeze Failed: " + e.message);
  }
}

function clearPreCampaignBaseline() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const baselineSheet = ss.getSheetByName(VDM_CONFIG.TABS.BASELINE);
    if (baselineSheet) ss.deleteSheet(baselineSheet);
    ui.alert("Baseline Snapshot Reset to Dynamic Live Mode.");
  } catch (e) {
    ui.alert("Baseline Reset Failed: " + e.message);
  }
}