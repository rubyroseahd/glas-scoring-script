/**
 * MODULE 2: INGESTION ENGINE
 * Extraction, Filtering, and Alphanumeric Sanitization Gateway
 */

function runDataIngestion() {
  try {
    const folder = DriveApp.getFolderById(VDM_CONFIG.FOLDER_ID);
    
    processShopifyFile(folder);
    processEEIFile(folder, VDM_CONFIG.SOURCE_FILES.EEI_USA, VDM_CONFIG.TABS.RAW_EEI_USA);
    processEEIFile(folder, VDM_CONFIG.SOURCE_FILES.EEI_WEB, VDM_CONFIG.TABS.RAW_EEI_WEB);
    processGenericCSV(folder, VDM_CONFIG.SOURCE_FILES.SALES, VDM_CONFIG.TABS.RAW_SALES);
    processGenericCSV(folder, VDM_CONFIG.SOURCE_FILES.COST, VDM_CONFIG.TABS.RAW_COST);
    
    resolveCostHierarchy();
    
  } catch (e) {
    logError("Ingestion", e);
    throw e;
  }
}

function processShopifyFile(folder) {
  const files = folder.getFilesByName(VDM_CONFIG.SOURCE_FILES.SHOPIFY);
  if (!files.hasNext()) {
    logError("Ingestion.processShopifyFile", `Shopify file not found: ${VDM_CONFIG.SOURCE_FILES.SHOPIFY}`);
    throw new Error(`Required file not found: ${VDM_CONFIG.SOURCE_FILES.SHOPIFY}. Please ensure it exists in the Drive folder.`);
  }
  
  const csvData = Utilities.parseCsv(files.next().getBlob().getDataAsString());
  const headers = csvData[0];
  const sIdx = getHeaderMap(headers);
  
  const processedMap = new Map();
  
  for (let i = 1; i < csvData.length; i++) {
    const row = csvData[i];
    const sku = sanitizeKey(row[sIdx["Variant SKU"]]);
    const status = (row[sIdx["Status"]] || "").toLowerCase();
    
    if (sku && status === "active" && !processedMap.has(sku)) {
      processedMap.set(sku, [sku, ...row]); // Prepend SKU Anchor
    }
  }
  
  const output = [["SKU_ANCHOR", ...headers], ...Array.from(processedMap.values())];
  const sheet = getOrCreateSheet(VDM_CONFIG.TABS.RAW_SHOPIFY, true);
  sheet.clear().getRange(1, 1, output.length, output[0].length).setValues(output);
  sheet.getRange(1, 1, output.length, 1).setNumberFormat("@"); // Format Anchor Column
}

function processEEIFile(folder, fileName, tabName) {
  const files = folder.getFilesByName(fileName);
  if (!files.hasNext()) {
    logError("Ingestion.processEEIFile", `EEI file not found: ${fileName}`);
    throw new Error(`Required file not found: ${fileName}. Please ensure it exists in the Drive folder.`);
  }
  
  const csvData = Utilities.parseCsv(files.next().getBlob().getDataAsString());
  if (csvData.length < 5) { // Need at least 5 rows for headers at row 5
    logError("Ingestion.processEEIFile", `EEI CSV is too short or empty: ${fileName}`);
    throw new Error(`EEI CSV is too short or empty: ${fileName}. Expected headers on row 5.`);
  }
  
  const headers = csvData[4];
  const eeiIdx = getHeaderMap(headers);
  
  const rows = csvData.slice(5).map(row => {
    const sku = sanitizeKey(row[eeiIdx["Item Code"]]);
    return [sku, ...row]; // Prepend SKU Anchor
  });
  
  const output = [["SKU_ANCHOR", ...headers], ...rows];
  const sheet = getOrCreateSheet(tabName, true);
  sheet.clear().getRange(1, 1, output.length, output[0].length).setValues(output);
  sheet.getRange(1, 1, output.length, 1).setNumberFormat("@");
}

function processGenericCSV(folder, fileName, tabName) {
  const files = folder.getFilesByName(fileName);
  if (!files.hasNext()) {
    logError("Ingestion.processGenericCSV", `Generic CSV file not found: ${fileName}`);
    throw new Error(`Required file not found: ${fileName}. Please ensure it exists in the Drive folder.`);
  }
  
  const data = Utilities.parseCsv(files.next().getBlob().getDataAsString());
  const headers = data[0];
  const skuColIdx = headers.findIndex(h => h.toLowerCase().includes("sku") || h === "Product variant SKU");
  
  const rows = data.slice(1).map(row => {
    const sku = sanitizeKey(row[skuColIdx]);
    return [sku, ...row];
  });
  
  const output = [["SKU_ANCHOR", ...headers], ...rows];
  const sheet = getOrCreateSheet(tabName, true);
  sheet.clear().getRange(1, 1, output.length, output[0].length).setValues(output);
  sheet.getRange(1, 1, output.length, 1).setNumberFormat("@");
}

function resolveCostHierarchy() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shopifySheet = ss.getSheetByName(VDM_CONFIG.TABS.RAW_SHOPIFY);
  const costSheet = ss.getSheetByName(VDM_CONFIG.TABS.RAW_COST);
  
  const shopifyData = shopifySheet.getDataRange().getValues();
  const costData = costSheet.getDataRange().getValues();
  
  const sIdx = getHeaderMap(shopifyData[0]);
  const cIdx = getHeaderMap(costData[0]);
  
  const costLookup = {};
  costData.slice(1).forEach(r => {
    costLookup[sanitizeKey(r[cIdx["SKU"]])] = {
      eei: parseFloat(r[cIdx["EEI LAST PURCHASE PRICE"]]) || 0,
      glas: parseFloat(r[cIdx["GLAS Costing"]]) || 0,
      cotr: parseFloat(r[cIdx["COTR LAST PURCHASE PRICE"]]) || 0
    };
  });
  
  const resolvedCosts = [["SKU Anchor", "Resolved Cost"]];
  
  shopifyData.slice(1).forEach(row => {
    const sku = sanitizeKey(row[sIdx["Variant SKU"]]);
    const shopifyCost = parseFloat(row[sIdx["Cost per item"]]) || 0;
    const external = costLookup[sku] || {eei: 0, glas: 0, cotr: 0};
    
    let finalCost = 0;
    if (external.eei > 0) finalCost = external.eei;
    else if (external.glas > 0) finalCost = external.glas;
    else if (external.cotr > 0) finalCost = external.cotr;
    else finalCost = shopifyCost;
    
    resolvedCosts.push([sku, finalCost]);
  });
  
  const masterCostSheet = getOrCreateSheet(VDM_CONFIG.TABS.MASTER_COST, true);
  masterCostSheet.clear().getRange(1, 1, resolvedCosts.length, 2).setValues(resolvedCosts);
}