/**
 * MODULE 4: LIGHTWEIGHT REPORTING WRAPPERS
 */

function recoverDashboardState() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName(VDM_CONFIG.TABS.DASHBOARD);
  if (!dashboard || dashboard.getLastRow() === 0) {
    return { headers: [], rows: [] };
  }
  const values = dashboard.getDataRange().getValues();
  return {
    headers: values[0] || [],
    rows: values.slice(1)
  };
}

function workflowRefreshActionHubOnly() {
  const ui = SpreadsheetApp.getUi();
  try {
    const state = recoverDashboardState();
    writeActionHubFromDashboardState(state);
    ui.alert("Action Items Hub refreshed from [02] Dashboard Matrix.");
  } catch (e) {
    ui.alert("Action Hub refresh failed: " + e.message);
  }
}

function workflowRefreshTierSummaryOnly() {
  const ui = SpreadsheetApp.getUi();
  try {
    const state = recoverDashboardState();
    writeTierSummaryFromDashboardState(state);
    ui.alert("Tier Summary refreshed from [02] Dashboard Matrix.");
  } catch (e) {
    ui.alert("Tier Summary refresh failed: " + e.message);
  }
}
