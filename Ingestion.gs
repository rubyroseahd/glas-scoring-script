/**
 * MODULE 2: INGESTION ENGINE
 */

function runDataIngestion() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const folder = DriveApp.getFolderById(VDM_CONFIG.FOLDER_ID);

    validateHeaders(folder);
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

function validateHeaders(folder) {
  validateShopifyHeaders(folder);
  validateSalesHeaders(folder);
  validateWarehouseHeaders(folder, VDM_CONFIG.SOURCE_FILES.EEI_USA);
  validateWarehouseHeaders(folder, VDM_CONFIG.SOURCE_FILES.EEI_WEB);
  validateCostHeaders(folder);
}

function ingestShopify(folder, ss) {
  const data = loadCsvFile(folder, VDM_CONFIG.SOURCE_FILES.SHOPIFY);
  const headers = data[0].map(h => safeStr(h));
  const hMap = getHeaderMap(headers);
  const skuHeader = getFirstAvailableHeader(hMap, ["VARIANT SKU", "SKU"]);
  const statusHeader = getFirstAvailableHeader(hMap, ["STATUS"]);
  const typeHeader = findFirstAvailableHeader(hMap, ["TYPE", "PRODUCT TYPE"]);

  const rows = [];
  const seen = new Set();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const sku = safeStr(row[hMap[skuHeader]]).toUpperCase();
    const status = safeStr(row[hMap[statusHeader]]).toLowerCase();
    if (!sku || status !== "active" || seen.has(sku)) continue;
    seen.add(sku);

    const processed = [sku];
    headers.forEach(h => {
      const idx = hMap[safeStr(h).toUpperCase()];
      const val = idx === undefined ? null : row[idx];
      processed.push(/price|qty|cost/i.test(h) ? safeNum(val) : safeStr(val));
    });
    const typeValue = typeHeader ? safeStr(row[hMap[typeHeader]]).toUpperCase() : "";
    processed.push(typeValue.includes("WEBONLY") ? "WEBONLY" : "SHARED");
    rows.push(processed);
  }

  writeToWorkbookTab(VDM_CONFIG.TABS.RAW_SHOPIFY, [["SKU_ANCHOR", ...headers, "FULFILLMENT TYPE"], ...rows], ss);
}

function ingestEEI(folder, fileName, tabName, ss) {
  const data = loadCsvFile(folder, fileName);
  if (data.length < 6) {
    throw new Error(`File ${fileName} contains insufficient rows (${data.length} found). Expected metadata rows 1–4, header on row 5, and data on row 6+.`);
  }
  const hMap = getHeaderMap(data[4]);
  const skuHeader = getFirstAvailableHeader(hMap, ["ITEM CODE", "SKU"]);
  const qtyHeader = getFirstAvailableHeader(
    hMap,
    tabName === VDM_CONFIG.TABS.RAW_EEI_USA
      ? ["EEI USA WAREHOUSE ON HAND STOCK", "QTY", "QUANTITY", "ON HAND STOCK"]
      : ["EEI WEB WAREHOUSE ON HAND STOCK", "QTY", "QUANTITY", "ON HAND STOCK"]
  );
  const salesHeader = findFirstAvailableHeader(hMap, ["SALES PAST 30 DAYS"]);
  const stockHeader = tabName === VDM_CONFIG.TABS.RAW_EEI_USA
    ? "EEI USA WAREHOUSE ON HAND STOCK"
    : "EEI WEB WAREHOUSE ON HAND STOCK";
  const rows = data.slice(5).map(r => {
    const sku = safeStr(r[hMap[skuHeader]]).toUpperCase();
    return [
      sku,
      sku,
      safeNum(r[hMap[qtyHeader]]) || 0,
      salesHeader ? safeNum(r[hMap[salesHeader]]) || 0 : 0
    ];
  }).filter(r => r[0] !== "");
  writeToWorkbookTab(tabName, [["SKU_ANCHOR", "ITEM CODE", stockHeader, "SALES PAST 30 DAYS"], ...rows], ss);
}

function ingestGenericCSV(folder, fileName, tabName, skuHeader, ss) {
  const data = loadCsvFile(folder, fileName);
  const headers = data[0].map(h => safeStr(h));
  const hMap = getHeaderMap(headers);
  if (tabName === VDM_CONFIG.TABS.RAW_COST) {
    getFirstAvailableHeader(hMap, ["SKU", "VARIANT SKU"]);
    getFirstAvailableHeader(hMap, ["LAST PURCHASE PRICE REPORT EEI(UPDATE)", "GLAS COSTING", "LAST PURCHASE PRICE REPORT COTR(UPDATE)", "COST", "UNIT COST", "COST PER ITEM"]);
  }
  const resolvedSkuHeader = findFirstAvailableHeader(hMap, [skuHeader, "VARIANT SKU", "SKU"]) || safeStr(skuHeader).toUpperCase();
  const skuIdx = hMap[resolvedSkuHeader];

  const rows = data.slice(1).map(r => {
    const sku = safeStr(r[skuIdx]).toUpperCase();
    if (!sku) return null;
    return [sku, ...r];
  }).filter(Boolean);

  writeToWorkbookTab(tabName, [["SKU_ANCHOR", ...headers], ...rows], ss);
}

function ingestSalesCSV(folder, ss) {
  const data = loadCsvFile(folder, VDM_CONFIG.SOURCE_FILES.SALES);
  const hMap = getHeaderMap(data[0]);
  const skuHeader = getFirstAvailableHeader(hMap, ["PRODUCT VARIANT SKU", "VARIANT SKU", "SKU"]);
  const qtyHeader = getFirstAvailableHeader(hMap, ["NET QUANTITY", "NET ITEMS SOLD", "QTY", "QUANTITY"]);

  const rows = data.slice(1).map(r => {
    const sku = safeStr(r[hMap[skuHeader]]).toUpperCase();
    return [sku, sku, safeNum(r[hMap[qtyHeader]]) || 0];
  }).filter(r => r[0] !== "");

  writeToWorkbookTab(VDM_CONFIG.TABS.RAW_SALES, [["SKU_ANCHOR", "Product variant SKU", "Net items sold"], ...rows], ss);
}

function executeCostResolutionWaterfall() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shopifySheet = ss.getSheetByName(VDM_CONFIG.TABS.RAW_SHOPIFY);
  const costSheet = ss.getSheetByName(VDM_CONFIG.TABS.RAW_COST);
  if (!shopifySheet || !costSheet) throw new Error("Required ingestion tabs are missing for cost resolution.");

  const shopifyData = shopifySheet.getDataRange().getValues();
  const costData = costSheet.getDataRange().getValues();
  const sIdxMap = getHeaderMap(shopifyData[0]);
  const cIdxMap = getHeaderMap(costData[0]);
  const shopifySkuHeader = getFirstAvailableHeader(sIdxMap, ["SKU_ANCHOR", "VARIANT SKU", "SKU"]);
  const costSkuHeader = getFirstAvailableHeader(cIdxMap, ["SKU_ANCHOR", "SKU ANCHOR", "SKU", "VARIANT SKU"]);
  const shopifyCostHeader = findFirstAvailableHeader(sIdxMap, ["COST PER ITEM"]);

  const cIdx = {
    sku: cIdxMap[costSkuHeader],
    eeiUpdate: cIdxMap["LAST PURCHASE PRICE REPORT EEI(UPDATE)"],
    glas: cIdxMap["GLAS COSTING"],
    cotrUpdate: cIdxMap["LAST PURCHASE PRICE REPORT COTR(UPDATE)"],
    cost: cIdxMap["COST"],
    unitCost: cIdxMap["UNIT COST"],
    costPerItem: cIdxMap["COST PER ITEM"]
  };

  const costMap = new Map();
  costData.slice(1).forEach(r => {
    const sku = safeStr(cIdx.sku === undefined ? "" : r[cIdx.sku]).toUpperCase();
    if (sku) costMap.set(sku, r);
  });

  const getNumericAt = (row, idx) => idx === undefined ? null : safeNum(row[idx]);

  const resolved = [["SKU Anchor", "Resolved Cost"]];
  shopifyData.slice(1).forEach(r => {
    const sku = safeStr(r[sIdxMap[shopifySkuHeader]]).toUpperCase();
    if (!sku) return;
    const shopifyCost = shopifyCostHeader ? getNumericAt(r, sIdxMap[shopifyCostHeader]) : null;
    const ext = costMap.get(sku);
    const final = resolveWaterfallCost(ext, cIdx, shopifyCost);
    resolved.push([sku, final]);
  });

  writeToWorkbookTab(VDM_CONFIG.TABS.MASTER_COST, resolved, ss);
}

function resolveWaterfallCost(costRow, costIndexes, shopifyCost) {
  const getNumericAt = (row, idx) => {
    if (!row || idx === undefined) return null;
    return safeNum(row[idx]);
  };
  const candidates = [
    getNumericAt(costRow, costIndexes.eeiUpdate),
    getNumericAt(costRow, costIndexes.glas),
    getNumericAt(costRow, costIndexes.cotrUpdate),
    getNumericAt(costRow, costIndexes.cost),
    getNumericAt(costRow, costIndexes.unitCost),
    getNumericAt(costRow, costIndexes.costPerItem),
    shopifyCost
  ];
  for (let i = 0; i < candidates.length; i++) {
    const value = candidates[i];
    if (typeof value === "number" && !isNaN(value) && value > 0) return value;
  }
  return 0;
}

function writeToWorkbookTab(name, data, ss) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  sheet.clear().clearFormats();
  if (data.length > 0 && data[0] && data[0].length > 0) {
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    sheet.getRange(1, 1, data.length, 1).setNumberFormat("@");
  }
}

function loadCsvFile(folder, fileName) {
  const files = folder.getFilesByName(fileName);
  if (!files.hasNext()) throw new Error(`Missing required file: ${fileName}`);
  const data = Utilities.parseCsv(files.next().getBlob().getDataAsString());
  if (!data || data.length === 0) throw new Error(`File ${fileName} is empty or malformed.`);
  return data;
}

function findFirstAvailableHeader(headerMap, candidates) {
  for (let i = 0; i < candidates.length; i++) {
    const normalized = safeStr(candidates[i]).toUpperCase();
    if (normalized && headerMap[normalized] !== undefined) return normalized;
  }
  return null;
}

function getFirstAvailableHeader(headerMap, candidates) {
  const header = findFirstAvailableHeader(headerMap, candidates);
  if (!header) throw new Error(`REQUIRED HEADER MISSING: expected one of ${candidates.join(", ")}`);
  return header;
}

function validateShopifyHeaders(folder) {
  const data = loadCsvFile(folder, VDM_CONFIG.SOURCE_FILES.SHOPIFY);
  const headers = getHeaderMap(data[0]);
  getFirstAvailableHeader(headers, ["VARIANT SKU", "SKU"]);
  getFirstAvailableHeader(headers, ["STATUS"]);
  getFirstAvailableHeader(headers, ["VARIANT PRICE", "PRICE"]);
}

function validateSalesHeaders(folder) {
  const data = loadCsvFile(folder, VDM_CONFIG.SOURCE_FILES.SALES);
  const headers = getHeaderMap(data[0]);
  getFirstAvailableHeader(headers, ["PRODUCT VARIANT SKU", "VARIANT SKU", "SKU"]);
  getFirstAvailableHeader(headers, ["NET QUANTITY", "NET ITEMS SOLD", "QTY", "QUANTITY"]);
}

function validateWarehouseHeaders(folder, fileName) {
  const data = loadCsvFile(folder, fileName);
  if (data.length < 6) {
    throw new Error(`File ${fileName} contains insufficient rows (${data.length} found). Expected metadata rows 1–4, header on row 5, and data on row 6+.`);
  }
  const headers = getHeaderMap(data[4]);
  getFirstAvailableHeader(headers, ["ITEM CODE", "SKU"]);
  getFirstAvailableHeader(
    headers,
    fileName === VDM_CONFIG.SOURCE_FILES.EEI_USA
      ? ["EEI USA WAREHOUSE ON HAND STOCK", "QTY", "QUANTITY", "ON HAND STOCK"]
      : ["EEI WEB WAREHOUSE ON HAND STOCK", "QTY", "QUANTITY", "ON HAND STOCK"]
  );
}

function validateCostHeaders(folder) {
  const data = loadCsvFile(folder, VDM_CONFIG.SOURCE_FILES.COST);
  const headers = getHeaderMap(data[0]);
  getFirstAvailableHeader(headers, ["SKU", "VARIANT SKU"]);
  getFirstAvailableHeader(headers, ["LAST PURCHASE PRICE REPORT EEI(UPDATE)", "GLAS COSTING", "LAST PURCHASE PRICE REPORT COTR(UPDATE)", "COST", "UNIT COST", "COST PER ITEM"]);
}
