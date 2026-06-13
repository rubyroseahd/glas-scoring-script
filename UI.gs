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