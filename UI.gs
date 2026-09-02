/**
 * MODULE 5: UI INTERFACE
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("EEI Pricing Engine Launcher")
    .addItem("1. Run Full System Pricing & Sync", "workflowRunFullSystemPricingAndSync")
    .addItem("1b. Refresh Dashboard Only (No Ingestion)", "executeDashboardRefresh")
    .addSeparator()
    .addItem("2. Refresh Action Items Hub Only", "workflowRefreshActionHubOnly")
    .addItem("3. Refresh Tier Summary Only", "workflowRefreshTierSummaryOnly")
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu("Advanced Diagnostics")
        .addItem("Run Pre-Flight Sanity Check", "workflowRunPreFlightSanityCheck")
        .addItem("Emergency Matrix Rollback", "workflowEmergencyMatrixRollback")
        .addItem("Purge Deprecated / Legacy Tabs", "workflowPurgeLegacyTabs")
        .addItem("Freeze Pre-Campaign Baseline Snapshot", "workflowFreezeBaselineSnapshot")
    )
    .addToUi();
}

function workflowRunFullSystemPricingAndSync() {
  const ui = SpreadsheetApp.getUi();
  try {
    runDataIngestion();
  } catch (e) {
    ui.alert(
      "Full sync stopped during ingestion. Dashboard refresh was skipped.\n\n" +
      "Fix source-file/header issues and rerun. Details (UI.gs::workflowRunFullSystemPricingAndSync -> Ingestion.gs::runDataIngestion): " + e.message
    );
    return { ok: false, stage: "ingestion", error: e.message };
  }

  try {
    const refreshResult = executeDashboardRefresh();
    ui.alert(
      "Full system pricing + sync completed.\n\n" +
      `Active SKUs processed: ${refreshResult.stats.totalActiveSkus}`
    );
    return { ok: true, stage: "refresh", result: refreshResult };
  } catch (e) {
    ui.alert(
      "Ingestion completed, but pricing refresh failed.\n\n" +
      "Review dashboard dependencies and rerun refresh. Details (UI.gs::workflowRunFullSystemPricingAndSync -> MatrixEngine.gs::executeDashboardRefresh): " + e.message
    );
    throw e;
  }
}

function workflowRunPreFlightSanityCheck() {
  const ui = SpreadsheetApp.getUi();
  try {
    const folder = DriveApp.getFolderById(getOperationalFolderId());
    validateHeaders(folder);
    ui.alert("Pre-Flight Sanity Check passed. Source CSV headers look valid.");
    return true;
  } catch (e) {
    ui.alert("Pre-Flight Sanity Check failed (UI.gs::workflowRunPreFlightSanityCheck -> Ingestion.gs::validateHeaders): " + e.message);
    return false;
  }
}

function workflowEmergencyMatrixRollback() {
  const ui = SpreadsheetApp.getUi();
  const choice = ui.alert(
    "Emergency Matrix Rollback",
    "Restore [02] Dashboard Matrix and dependent outputs from _backup_matrix_data?\n\nThis will overwrite current run outputs.",
    ui.ButtonSet.YES_NO
  );
  if (choice !== ui.Button.YES) {
    ui.alert("Emergency Matrix Rollback cancelled.");
    return false;
  }

  try {
    const state = readBackupMatrixState_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    writeTableToSheet_(ss, VDM_CONFIG.TABS.DASHBOARD, state.headers, state.rows);
    writeTableToSheet_(ss, VDM_CONFIG.TABS.BI_FEED, state.headers, state.rows);
    writeActionHubFromDashboardState(state);
    writeTierSummaryFromDashboardState(state);
    writeSyncAuditFromDashboardState(state);
    writeMasterLedgerFromDashboardState(state);
    ui.alert("Emergency Matrix Rollback complete. Dashboard, queues, and audit outputs were restored.");
    return true;
  } catch (e) {
    ui.alert("Emergency Matrix Rollback failed (UI.gs::workflowEmergencyMatrixRollback): " + e.message);
    return false;
  }
}

function workflowPurgeLegacyTabs() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const purgedNames = getLegacyTabNames_(ss.getSheets().map(sheet => sheet.getName()));
  if (purgedNames.length === 0) {
    ui.alert("No Legacy Tabs", "All tabs conform to the active VDM schema.", ui.ButtonSet.OK);
    return { count: 0, names: [] };
  }
  const response = ui.alert(
    "Confirm Tab Purge",
    "The following " + purgedNames.length + " tab(s) will be permanently deleted:\n\n" +
      purgedNames.join("\n") +
      "\n\nDo you want to proceed?",
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return { count: 0, names: [] };
  const result = purgeLegacyTabs();
  const purgedCount = result.count;
  const deletedNames = result.names;
  ui.alert("Cleanup Complete", "Purged " + purgedCount + " legacy tabs:\n" + deletedNames.join(", "), ui.ButtonSet.OK);
  return result;
}

function readBackupMatrixState_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const backup = ss.getSheetByName(VDM_CONFIG.TABS.BACKUP_MATRIX_DATA);
  if (!backup || backup.getLastRow() < 2) {
    throw new Error("_backup_matrix_data is missing or empty.");
  }
  const values = backup.getDataRange().getValues();
  return {
    headers: values[0] || [],
    rows: values.slice(1)
  };
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
    ui.alert("Baseline snapshot failed (UI.gs::workflowFreezeBaselineSnapshot): " + e.message);
  }
}
