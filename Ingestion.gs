/**
 * MODULE 2: INGESTION ENGINE
 */

function runDataIngestion() {
  try {
    const folder = DriveApp.getFolderById(VDM_CONFIG.FOLDER_ID);
    
    ingestShopify(folder);
    ingestEEI(folder, VDM_CONFIG.SOURCE_FILES.EEI_USA, VDM_CONFIG.TABS.RAW_EEI_USA);
    ingestEEI(folder, VDM_CONFIG.SOURCE_FILES.EEI_WEB, VDM_CONFIG.TABS.RAW_EEI_WEB);
    ingestGenericCSV(folder, VDM_CONFIG.SOURCE_FILES.SALES, VDM_CONFIG.TABS.RAW_SALES, "Product variant SKU");
    ingestGenericCSV(folder, VDM_CONFIG.SOURCE_FILES.COST, VDM_CONFIG.TABS.RAW_COST, "SKU");
    
    executeCostResolutionWaterfall();
  } catch (e) {
    logError("Ingestion", e);
    throw e;
  }
}

function sanitize(val) {
  if (val === null || val === undefined) return "";
  return String(val).trim().toUpperCase().replace(/[\r\n\t]+/g, "");
}

function ingestShopify(folder) {
  const files = folder.getFilesByName(VDM_CONFIG.SOURCE_FILES.SHOPIFY);
  if (!files.hasNext()) throw new Error("Shopify file missing");
  
  const data = Utilities.parseCsv(files.next().getBlob().getDataAsString());
  const headers = data[0];
  const skuIdx = headers.indexOf("Variant SKU");
  const statusIdx = headers.indexOf("Status");
  const priceIdx = headers.indexOf("Variant Price");
  const compareIdx = headers.indexOf("Variant Compare At Price");
  const vendorIdx = headers.indexOf("Vendor");
  const fulfillIdx = headers.indexOf("Fulfillment service");
  const qtyIdx = headers.indexOf("Variant Inventory Qty");
  const costIdx = headers.indexOf("Cost per item");

  const map = new Map();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const sku = sanitize(row[skuIdx]);
    if (sku && row[statusIdx].toLowerCase() === "active" && !map.has(sku)) {
      // Structure: Anchor, SKU, Status, Price, Compare, Fulfill, Vendor, Qty, ShopifyCost
      map.set(sku, [sku, sku, row[statusIdx], row[priceIdx], row[compareIdx], row[fulfillIdx], row[vendorIdx], row[qtyIdx], row[costIdx]]);
    }
  }

  const outHeaders = ["SKU_ANCHOR", "Variant SKU", "Status", "Variant Price", "Variant Compare At Price", "Fulfillment service", "Vendor", "Variant Inventory Qty", "Cost per item"];
  writeToHiddenTab(VDM_CONFIG.TABS.RAW_SHOPIFY, [outHeaders, ...Array.from(map.values())]);
}

function ingestEEI(folder, fileName, tabName) {
  const files = folder.getFilesByName(fileName);
  if (!files.hasNext()) throw new Error(`${fileName} missing`);
  
  const data = Utilities.parseCsv(files.next().getBlob().getDataAsString());
  const headers = data[4]; // Row 5 calibration
  const skuIdx = headers.indexOf("Item Code");
  const stockIdx = 11; // Column L
  const salesIdx = headers.indexOf("Sales Past 30 Days");

  const rows = data.slice(5).map(r => {
    const sku = sanitize(r[skuIdx]);
    return [sku, sku, r[stockIdx], r[salesIdx] || 0];
  });

  writeToHiddenTab(tabName, [["SKU_ANCHOR", "Item Code", "On Hand", "Sales 30D"], ...rows]);
}

function ingestGenericCSV(folder, fileName, tabName, skuHeader) {
  const files = folder.getFilesByName(fileName);
  if (!files.hasNext()) throw new Error(`${fileName} missing`);
  
  const data = Utilities.parseCsv(files.next().getBlob().getDataAsString());
  const headers = data[0];
  const skuIdx = headers.indexOf(skuHeader);
  const netSalesIdx = headers.indexOf("Net items sold");

  const rows = data.slice(1).map(r => {
    const sku = sanitize(r[skuIdx]);
    // If sales data, ensure we extract specific columns. Otherwise keep row.
    return [sku, ...r];
  });

  writeToHiddenTab(tabName, [["SKU_ANCHOR", ...headers], ...rows]);
}

function executeCostResolutionWaterfall() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shopifyData = ss.getSheetByName(VDM_CONFIG.TABS.RAW_SHOPIFY).getDataRange().getValues();
  const costData = ss.getSheetByName(VDM_CONFIG.TABS.RAW_COST).getDataRange().getValues();
  
  const costHeaders = costData[0];
  const cIdx = {
    sku: costHeaders.indexOf("SKU"),
    eei: costHeaders.indexOf("EEI LAST PURCHASE PRICE"),
    glas: costHeaders.indexOf("GLAS Costing"),
    cotr: costHeaders.indexOf("COTR LAST PURCHASE PRICE")
  };

  const costMap = new Map();
  costData.slice(1).forEach(r => costMap.set(sanitize(r[cIdx.sku]), r));

  const resolved = [["SKU Anchor", "Resolved Cost"]];
  shopifyData.slice(1).forEach(r => {
    const sku = r[0];
    const shopifyCost = parseFloat(r[8]) || 0;
    const ext = costMap.get(sku);
    
    let final = 0;
    if (ext) {
      final = parseFloat(ext[cIdx.eei]) || parseFloat(ext[cIdx.glas]) || parseFloat(ext[cIdx.cotr]) || shopifyCost;
    } else {
      final = shopifyCost;
    }
    resolved.push([sku, final]);
  });

  writeToHiddenTab(VDM_CONFIG.TABS.MASTER_COST, resolved);
}

function writeToHiddenTab(name, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.hideSheet();
  }
  sheet.clear().clearFormats();
  if (data.length > 0) {
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    sheet.getRange(1, 1, data.length, 1).setNumberFormat("@");
  }
}

function applyHeaderStyle(range) {
  range.setBackground(VDM_CONFIG.DESIGN.HEADER_BG)
       .setFontColor(VDM_CONFIG.DESIGN.HEADER_TEXT)
       .setFontWeight("bold")
       .setHorizontalAlignment("center");
}