/**
 * MODULE 5: UI INTERFACE & SYSTEM SAFEGUARDS
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('VDM Engine Launcher')
    .addItem('1. Build Master Architecture', 'triggerNuclearArchitectureWipe')
    .addSeparator()
    .addItem('2. Run VDM Data Refresh', 'runFullRefreshCycle')
    .addToUi();
}

function triggerNuclearArchitectureWipe() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'CRITICAL WARNING',
    'This will wipe all data across all tabs and rebuild the system architecture. This cannot be undone. Proceed?',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tabKeys = Object.values(VDM_CONFIG.TABS);
    
    tabKeys.forEach(name => {
      let sheet = ss.getSheetByName(name);
      if (sheet) {
        sheet.clear();
        sheet.clearConditionalFormatRules();
      } else {
        sheet = ss.insertSheet(name);
      }
      
      if (name.startsWith('_')) {
        sheet.hideSheet();
      }
    });
    
    try {
      // Initialize Control Panel Headers
      const cp = ss.getSheetByName(VDM_CONFIG.TABS.CONTROL);
      const headerRange = cp.getRange("A1:E1");
      headerRange.setValues([["Active GWP SKUs", "New Launch Overrides", "MAP Restricted Brands", "", "Affiliate Stack Rate"]]);
      applyHeaderStyle(headerRange); // Fix 3: Use applyHeaderStyle
      cp.getRange("E2").setValue(0.15).setNumberFormat("0%"); // Default stack rate
      
      ui.alert("Architecture rebuilt successfully.");
    } catch (e) {
      logError("UI-Setup", e);
    }
  }
}