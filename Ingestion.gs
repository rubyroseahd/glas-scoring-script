/**
 * MODULE 2: INGESTION ENGINE
 */

function runDataIngestion() {
  try {
    const folder = DriveApp.getFolderById(VDM_CONFIG.FOLDER_ID);
    
    ingestShopify(folder);
    ingestEEI(folder, VDM_CONFIG.SOURCE_FILES.EEI_USA, VDM_CONFIG.TABS.RAW_EEI_USA);
    ingestEEI(folder, VDM_CONFIG.SOURCE_FILES.EEI_WEB, VDM_CONFIG.TABS.RAW_EEI_WEB);
    ingestSalesCSV(folder);
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
  const hMap = getHeaderMap(data[0]);
  const targetHeaders = VDM_CONFIG.HEADERS.SHOPIFY;

  const map = new Map();
  for (var i = 1; i < data.length; i++) {
    const row = data[i];
    const sku = safeStr(row[hMap["Variant SKU"]]).toUpperCase();
    const status = safeStr(row[hMap["Status"]]).toLowerCase();
    
    if (sku && status === "active" && !map.has(sku)) {
      // Map row to the exact sequence in outHeaders
      const processedRow = [sku]; // Column A Anchor
      targetHeaders.forEach(h => {
        const val = row[hMap[h]];
        processedRow.push(h.includes("Price") || h.includes("Qty") || h.includes("item") ? safeNum(val) : safeStr(val));
      });
      map.set(sku, processedRow);
    }
  }

  const outHeaders = ["SKU_ANCHOR", ...targetHeaders];
  writeToHiddenTab(VDM_CONFIG.TABS.RAW_SHOPIFY, [outHeaders, ...Array.from(map.values())]);
}

function ingestEEI(folder, fileName, tabName) {
  const files = folder.getFilesByName(fileName);
  if (!files.hasNext()) throw new Error(`${fileName} missing`);
  
  const data = Utilities.parseCsv(files.next().getBlob().getDataAsString());
  const hRow = data[4]; 
  const hMap = getHeaderMap(hRow);
  const targetHeaders = tabName === VDM_CONFIG.TABS.RAW_EEI_USA ? VDM_CONFIG.HEADERS.USA_WAREHOUSE : VDM_CONFIG.HEADERS.WEB_WAREHOUSE;
  
  const rows = data.slice(5).map(r => {
    const sku = safeStr(r[hMap["Item Code"]]).toUpperCase();
    const out = [sku];
    targetHeaders.forEach(h => {
      const val = r[hMap[h]];
      out.push(h.includes("Stock") || h.includes("Sales") || h.includes("days") ? safeNum(val) : safeStr(val));
    });
    return out;
  });

  writeToHiddenTab(tabName, [["SKU_ANCHOR", ...targetHeaders], ...rows]);
}

function ingestGenericCSV(folder, fileName, tabName, skuHeader) {
  const files = folder.getFilesByName(fileName);
  if (!files.hasNext()) throw new Error(`${fileName} missing`);
  
  const data = Utilities.parseCsv(files.next().getBlob().getDataAsString());
  const headers = data[0];
  const skuIdx = headers.indexOf(skuHeader);

  const rows = data.slice(1).map(r => {
    const sku = safeStr(r[skuIdx]).toUpperCase();
    if (!sku) return null;
    // If sales data, ensure we extract specific columns. Otherwise keep row.
    return [sku, ...r];
  }).filter(r => r !== null);

  writeToHiddenTab(tabName, [["SKU_ANCHOR", ...headers], ...rows]);
}

/**
 * Dedicated Sales Ingestion to maintain strict 3-column contract: [SKU_ANCHOR, SKU, Net items sold]
 * Prevents downstream VLOOKUP index corruption.
 */
function ingestSalesCSV(folder) {
  const files = folder.getFilesByName(VDM_CONFIG.SOURCE_FILES.SALES);
  if (!files.hasNext()) throw new Error("Sales file missing");
  
  const data = Utilities.parseCsv(files.next().getBlob().getDataAsString());
  const hIdx = getHeaderMap(data[0]);
  const salesCol = hIdx["Net items sold"];
  
  const rows = data.slice(1).map(r => {
    const sku = safeStr(r[hIdx["Product variant SKU"]]).toUpperCase();
    return [sku, sku, safeNum(r[salesCol])];
  }).filter(r => r[0] !== "");

  writeToHiddenTab(VDM_CONFIG.TABS.RAW_SALES, [["SKU_ANCHOR", "Product variant SKU", "Net items sold"], ...rows]);
}

function executeCostResolutionWaterfall() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shopifySheet = ss.getSheetByName(VDM_CONFIG.TABS.RAW_SHOPIFY);
  const shopifyData = shopifySheet.getDataRange().getValues();
  const costData = ss.getSheetByName(VDM_CONFIG.TABS.RAW_COST).getDataRange().getValues();
  
  const shopifyHeaders = shopifyData[0];
  const costHeaders = costData[0];
  
  const sIdx = { cost: shopifyHeaders.indexOf("Cost per item") };
  const cIdx = {
    sku: costHeaders.indexOf("SKU"),
    eei: costHeaders.indexOf("EEI LAST PURCHASE PRICE"),
    glas: costHeaders.indexOf("GLAS Costing"),
    cotr: costHeaders.indexOf("COTR LAST PURCHASE PRICE")
  };

  const costMap = new Map();
  costData.slice(1).forEach(r => costMap.set(safeStr(r[cIdx.sku]).toUpperCase(), r));

  const resolved = [["SKU Anchor", "Resolved Cost"]];
  shopifyData.slice(1).forEach(r => {
    const sku = r[0];
    const shopifyCost = safeNum(r[sIdx.cost]);
    const ext = costMap.get(sku);
    
    let final = 0;
    if (ext) {
      final = safeNum(ext[cIdx.eei]) || safeNum(ext[cIdx.glas]) || safeNum(ext[cIdx.cotr]) || shopifyCost;
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
  if (data.length > 0 && data[0] && data[0].length > 0) {
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