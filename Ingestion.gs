/**
 * MODULE 2: INGESTION ENGINE
 */

function runDataIngestion() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const folder = DriveApp.getFolderById(VDM_CONFIG.FOLDER_ID);
    
    // Pre-Flight Header Check
    validateHeaders(folder);

    // Pass spreadsheet object to avoid repeated calls
    ingestShopify(folder, ss);
    ingestEEI(folder, VDM_CONFIG.SOURCE_FILES.EEI_USA, VDM_CONFIG.TABS.RAW_EEI_USA, ss);
    ingestEEI(folder, VDM_CONFIG.SOURCE_FILES.EEI_WEB, VDM_CONFIG.TABS.RAW_EEI_WEB, ss);
    ingestSalesCSV(folder, ss);
    ingestGenericCSV(folder, VDM_CONFIG.SOURCE_FILES.COST, VDM_CONFIG.TABS.RAW_COST, "SKU", ss);
    
    executeCostResolutionWaterfall();
  } catch (e) {
    logError("Ingestion", e);
    throw e;
  }
}

/**
 * Validates that all required headers exist in source CSVs before processing.
 * @param {GoogleAppsScript.Drive.Folder} folder
 */
function validateHeaders(folder) {
  const configs = [
    { file: VDM_CONFIG.SOURCE_FILES.SHOPIFY, headers: VDM_CONFIG.HEADERS.SHOPIFY, skip: 0 },
    { file: VDM_CONFIG.SOURCE_FILES.EEI_USA, headers: VDM_CONFIG.HEADERS.USA_WAREHOUSE, skip: 4 },
    { file: VDM_CONFIG.SOURCE_FILES.EEI_WEB, headers: VDM_CONFIG.HEADERS.WEB_WAREHOUSE, skip: 4 },
    { file: VDM_CONFIG.SOURCE_FILES.SALES, headers: VDM_CONFIG.HEADERS.RETAIL_VELOCITY, skip: 0 },
    { file: VDM_CONFIG.SOURCE_FILES.COST, headers: VDM_CONFIG.HEADERS.COST_WATERFALL, skip: 0 }
  ];

  configs.forEach(cfg => {
    const files = folder.getFilesByName(cfg.file);
    if (!files.hasNext()) throw new Error(`Missing required file: ${cfg.file}`);
    const data = Utilities.parseCsv(files.next().getBlob().getDataAsString()); // data is 0-indexed
    if (!data || data.length < cfg.skip + 1) throw new Error(`File ${cfg.file} is empty or malformed, or header row not found at expected index ${cfg.skip}.`);
    
    const fileHeaders = data[cfg.skip].map(h => h.toString().trim().toUpperCase());
    cfg.headers.forEach(req => {
      if (!fileHeaders.includes(req.toUpperCase())) {
        throw new Error(`File "${cfg.file}" is missing required column: "${req}"`);
      }
    });
  });
}

function ingestShopify(folder, ss) {
  const files = folder.getFilesByName(VDM_CONFIG.SOURCE_FILES.SHOPIFY);
  if (!files.hasNext()) throw new Error("Shopify file missing");
  
  const data = Utilities.parseCsv(files.next().getBlob().getDataAsString());
  const hMap = getHeaderMap(data[0]);
  const targetHeaders = VDM_CONFIG.HEADERS.SHOPIFY;

  const map = new Map();
  for (var i = 1; i < data.length; i++) {
    const row = data[i];
    const sku = safeStr(row[hMap[VDM_CONFIG.HEADERS.SHOPIFY[0]]]).toUpperCase();
    const status = safeStr(row[hMap[VDM_CONFIG.HEADERS.SHOPIFY[2]]]).toLowerCase();
    
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
  writeToHiddenTab(VDM_CONFIG.TABS.RAW_SHOPIFY, [outHeaders, ...Array.from(map.values())], ss);
}

function ingestEEI(folder, fileName, tabName, ss) {
  const files = folder.getFilesByName(fileName);
  if (!files.hasNext()) throw new Error(`${fileName} missing`);
  
  const data = Utilities.parseCsv(files.next().getBlob().getDataAsString());
  if (data.length <= 4) throw new Error(`File ${fileName} is too short or malformed.`);
  const hRow = data[4];
  const hMap = getHeaderMap(hRow);
  const targetHeaders = tabName === VDM_CONFIG.TABS.RAW_EEI_USA ? VDM_CONFIG.HEADERS.USA_WAREHOUSE : VDM_CONFIG.HEADERS.WEB_WAREHOUSE;
  
  const rows = data.slice(5).map(r => {
    const sku = safeStr(r[hMap[targetHeaders[0]]]).toUpperCase();
    const out = [sku];
    targetHeaders.forEach(h => {
      const val = r[hMap[h]];
      out.push(h.includes("Stock") || h.includes("Sales") || h.includes("days") ? safeNum(val) : safeStr(val));
    });
    return out;
  });

  writeToHiddenTab(tabName, [["SKU_ANCHOR", ...targetHeaders], ...rows], ss);
}

function ingestGenericCSV(folder, fileName, tabName, skuHeader, ss) {
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

  writeToHiddenTab(tabName, [["SKU_ANCHOR", ...headers], ...rows], ss);
}

/**
 * Dedicated Sales Ingestion to maintain strict 3-column contract: [SKU_ANCHOR, SKU, Net items sold]
 * Prevents downstream VLOOKUP index corruption.
 */
function ingestSalesCSV(folder, ss) {
  const files = folder.getFilesByName(VDM_CONFIG.SOURCE_FILES.SALES);
  if (!files.hasNext()) throw new Error("Sales file missing");
  
  const data = Utilities.parseCsv(files.next().getBlob().getDataAsString());
  const hIdx = getHeaderMap(data[0]);
  const salesCol = hIdx[VDM_CONFIG.HEADERS.RETAIL_VELOCITY[1]];
  
  const rows = data.slice(1).map(r => {
    const sku = safeStr(r[hIdx[VDM_CONFIG.HEADERS.RETAIL_VELOCITY[0]]]).toUpperCase();
    return [sku, sku, safeNum(r[salesCol])];
  }).filter(r => r[0] !== "");

  writeToHiddenTab(VDM_CONFIG.TABS.RAW_SALES, [["SKU_ANCHOR", "Product variant SKU", "Net items sold"], ...rows], ss);
}

function executeCostResolutionWaterfall() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shopifySheet = ss.getSheetByName(VDM_CONFIG.TABS.RAW_SHOPIFY);
  const shopifyData = shopifySheet.getDataRange().getValues();
  const costData = ss.getSheetByName(VDM_CONFIG.TABS.RAW_COST).getDataRange().getValues();
  
  const sIdxMap = getHeaderMap(shopifyData[0]);
  const cIdxMap = getHeaderMap(costData[0]);
  
  const sIdx = { cost: sIdxMap["COST PER ITEM"] };
  const cIdx = {
    sku: cIdxMap["SKU_ANCHOR"],
    eei: cIdxMap["EEI LAST PURCHASE PRICE"],
    glas: cIdxMap["GLAS COSTING"],
    cotr: cIdxMap["COTR LAST PURCHASE PRICE"]
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
      final = safeNum(ext[cIdx.eei]) || safeNum(ext[cIdx.glas]) || safeNum(ext[cIdx.cotr]) || shopifyCost || 0;
    } else {
      final = shopifyCost || 0;
    }
    resolved.push([sku, safeNum(final)]);
  });

  writeToHiddenTab(VDM_CONFIG.TABS.MASTER_COST, resolved, ss);
}

function writeToHiddenTab(name, data, ss) {
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