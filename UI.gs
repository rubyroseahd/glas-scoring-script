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
      .addItem('Commit Shopify Sync Only (Bypass Matrix)', 'commitShopifySyncOnly')
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
    const settingsHeaders = [["Active GWP SKUs", "New Launch Overrides", "MAP Restricted Brands", "B2B Reserve Min Qty", "Affiliate Coupon Rate"]];
    settings.getRange(1, 1, 1, 5).setValues(settingsHeaders);
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
    const folder = DriveApp.getFolderById(VDM_CONFIG.FOLDER_ID);
    validateHeaders(folder);
    ui.alert("SUCCESS: All source file headers and directory structures validated.");
  } catch (e) {
    ui.alert("SANITY CHECK FAILED: " + e.message);
  }
}

function commitShopifySyncOnly() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const state = recoverDashboardState();
    const shopifyMap = getShopifyMap();
    generateSyncAudit(ss, state.rows, getHeaderMap(state.headers), shopifyMap);
    ui.alert("Storefront Sync Refreshed (Matrix Engine Bypassed).");
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