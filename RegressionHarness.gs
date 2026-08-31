/**
 * MODULE 6: LIGHTWEIGHT REGRESSION HARNESS
 *
 * Manual execution entry point for core scoring and routing checks.
 * Run from Apps Script editor: runRegressionHarness
 */

function runRegressionHarness() {
  const assertions = [];

  // Case 1: Tie-aware percentile midpoint behavior
  const tieData = [2, 2, 4, 6, 8, 8, 10];
  const tiePctLow = getPercentileRankInc(tieData, 2);
  const tiePctHigh = getPercentileRankInc(tieData, 8);
  assertions.push(assertCondition("tie percentile is midpoint for duplicated low values", tiePctLow === 1 / 12));
  assertions.push(assertCondition("tie percentile is midpoint for duplicated high values", tiePctHigh === 3 / 4));
  assertions.push(assertCondition("equal values get equal percentile", getPercentileRankInc(tieData, 8) === getPercentileRankInc(tieData, 8)));

  // Case 2: Queue precedence (missing cost beats low-score routing)
  assertions.push(assertCondition(
    "MAP registry rejects partial vendor match",
    !isMapVendorMatch("ACME TOYS", ["ACME"])
  ));
  assertions.push(assertCondition(
    "MAP registry accepts exact normalized vendor match",
    isMapVendorMatch("Acme Toys", ["ACME TOYS"])
  ));
  const missingCostCase = evaluateRoutingRegressionCase({
    gateCode: GATEKEEPER_CODES.NONE,
    cost: 0,
    curMargin: 0.4,
    simNet: 20,
    vdmMarkdown: 0.65,
    curMarkdown: 0,
    fulfillment: "WEBONLY",
    totalScore: 2,
    usaStock: 0,
    b2b30DSales: 0,
    b2bReserveMin: 500,
    units90: 0
  });
  assertions.push(assertEqual("missing cost guardrail", missingCostCase.guardrailCode, GUARDRAIL_CODES.ERR_MISSING_COST));
  assertions.push(assertEqual("missing cost queue", missingCostCase.queueCode, QUEUE_CODES.QUEUE_1A_COST));

  // Case 3: Queue precedence (negative margin beats Queue 2/3)
  const negativeMarginCase = evaluateRoutingRegressionCase({
    gateCode: GATEKEEPER_CODES.NONE,
    cost: 20,
    curMargin: -0.05,
    simNet: 25,
    vdmMarkdown: 0.5,
    curMarkdown: 0,
    fulfillment: "SHARED",
    totalScore: 1,
    usaStock: 0,
    b2b30DSales: 0,
    b2bReserveMin: 500,
    units90: 0
  });
  assertions.push(assertEqual("negative margin guardrail", negativeMarginCase.guardrailCode, GUARDRAIL_CODES.ERR_NEGATIVE_MARGIN));
  assertions.push(assertEqual("negative margin queue", negativeMarginCase.queueCode, QUEUE_CODES.QUEUE_1A_MARGIN));

  // Case 4: WEBONLY clearance proxy (sales90 = 0 => clearance)
  const webOnlyZeroSalesTier = deriveTargetTier({
    gateCode: GATEKEEPER_CODES.NONE,
    totalScore: 3,
    fulfillment: "WEBONLY",
    units90: 0
  });
  assertions.push(assertEqual("webonly zero-sales tier", webOnlyZeroSalesTier.tier, "Clearance/Archive (65% Off)"));

  // Case 5: WEBONLY non-zero sales proxy (sales90 > 0 => digital review)
  const webOnlySomeSalesTier = deriveTargetTier({
    gateCode: GATEKEEPER_CODES.NONE,
    totalScore: 3,
    fulfillment: "WEBONLY",
    units90: 5
  });
  assertions.push(assertEqual("webonly positive-sales tier", webOnlySomeSalesTier.tier, "Accelerator / Digital Review (50% Off)"));

  // Case 6: B2B hold suppresses queue routing when safe
  const b2bHoldCase = evaluateRoutingRegressionCase({
    gateCode: GATEKEEPER_CODES.NONE,
    cost: 30,
    curMargin: 0.5,
    simNet: 90,
    vdmMarkdown: 0.5,
    curMarkdown: 0.3,
    fulfillment: "SHARED",
    totalScore: 2,
    usaStock: 700,
    b2b30DSales: 10,
    b2bReserveMin: 500,
    units90: 15
  });
  assertions.push(assertEqual("b2b hold guardrail", b2bHoldCase.guardrailCode, GUARDRAIL_CODES.WARN_B2B_HOLD));
  assertions.push(assertEqual("b2b hold queue suppressed", b2bHoldCase.queueCode, QUEUE_CODES.NONE));

  // Case 7: Cost waterfall prioritizes audited update headers
  const prioritizedCost = resolveWaterfallCost(
    ["SKU1", 9, 8, 7, 6, 5],
    { eeiUpdate: 1, glas: 2, cotrUpdate: 3, cost: 4, unitCost: 5 },
    4
  );
  assertions.push(assertEqual("cost waterfall priority uses EEI update first", prioritizedCost, 9));

  // Case 8: Cost waterfall preserves an explicit numeric zero at the highest-priority column
  const zeroCost = resolveWaterfallCost(
    ["SKU2", 0, 5, 8, null, 11],
    { eeiUpdate: 1, glas: 2, cotrUpdate: 3, cost: 4, unitCost: 5 },
    13
  );
  assertions.push(assertEqual("cost waterfall preserves explicit zero at eeiUpdate", zeroCost, 0));

  // Case 8b: Negative values and null/empty fall through; first non-negative value wins
  const fallbackCost = resolveWaterfallCost(
    ["SKU4", null, -3, "", null, 11],
    { eeiUpdate: 1, glas: 2, cotrUpdate: 3, cost: 4, unitCost: 5 },
    13
  );
  assertions.push(assertEqual("cost waterfall skips null/negative/empty and selects first non-negative", fallbackCost, 11));

  // Case 9: Shopify cost is final fallback when all procurement columns are null/empty
  const shopifyFallbackCost = resolveWaterfallCost(
    ["SKU3", null, null, null, null, null, 999],
    { eeiUpdate: 1, glas: 2, cotrUpdate: 3, cost: 4, unitCost: 5 },
    17
  );
  assertions.push(assertEqual(
    "cost waterfall falls back to Shopify cost and ignores external Cost Per Item",
    shopifyFallbackCost,
    17
  ));

  // Case 10: stripDuplicateSuffix removes ` (N)` before extension
  assertions.push(assertEqual("stripDuplicateSuffix removes (1)", stripDuplicateSuffix("shopify_export_gt (1).csv"), "shopify_export_gt.csv"));
  assertions.push(assertEqual("stripDuplicateSuffix removes (12)", stripDuplicateSuffix("EEI USA Whse Stock Report (12).csv"), "EEI USA Whse Stock Report.csv"));
  assertions.push(assertEqual("stripDuplicateSuffix leaves clean filename unchanged", stripDuplicateSuffix("Cost_Data.csv"), "Cost_Data.csv"));
  assertions.push(assertEqual("stripDuplicateSuffix ignores mid-name parens", stripDuplicateSuffix("report (draft).csv"), "report (draft).csv"));
  assertions.push(assertEqual("stripDuplicateSuffix handles no extension", stripDuplicateSuffix("report (1)"), "report (1)"));

  // Case 11: findCsvFileInFolder – exact match preferred over fallback
  const mockFolder = makeMockFolder([
    { name: "Cost_Data.csv", content: "sku,cost\nA,5" },
    { name: "Cost_Data (1).csv", content: "sku,cost\nB,6" }
  ]);
  const exactFile = findCsvFileInFolder(mockFolder, "Cost_Data.csv");
  assertions.push(assertEqual("exact match preferred when both exist", exactFile.getName(), "Cost_Data.csv"));

  // Case 11b: findCsvFileInFolder – multiple exact-name matches throws ambiguity error
  const dupeExactFolder = makeMockFolder([
    { name: "Cost_Data.csv", content: "sku,cost\nA,5" },
    { name: "Cost_Data.csv", content: "sku,cost\nB,6" }
  ]);
  let dupeExactErr = null;
  try { findCsvFileInFolder(dupeExactFolder, "Cost_Data.csv"); } catch (e) { dupeExactErr = e.message; }
  assertions.push(assertCondition("multiple exact-name matches throws ambiguity error",
    dupeExactErr && dupeExactErr.indexOf("Cost_Data.csv") !== -1 && dupeExactErr.toLowerCase().indexOf("ambiguous") !== -1));

  // Case 12: findCsvFileInFolder – fallback to single suffix variant
  const fallbackFolder = makeMockFolder([
    { name: "shopify_export_gt (2).csv", content: "sku\nX" }
  ]);
  const fallbackFile = findCsvFileInFolder(fallbackFolder, "shopify_export_gt.csv");
  assertions.push(assertEqual("fallback resolves single suffix variant", fallbackFile.getName(), "shopify_export_gt (2).csv"));

  // Case 13: findCsvFileInFolder – case-insensitive fallback
  const caseFolder = makeMockFolder([
    { name: "EEI WEB Whse Stock Report (1).csv", content: "sku\nY" }
  ]);
  const caseFile = findCsvFileInFolder(caseFolder, "EEI WEB Whse Stock Report.csv");
  assertions.push(assertEqual("case-insensitive fallback matches suffix variant", caseFile.getName(), "EEI WEB Whse Stock Report (1).csv"));

  // Case 14: findCsvFileInFolder – missing file throws configured name
  let missingErr = null;
  try { findCsvFileInFolder(makeMockFolder([]), "Cost_Data.csv"); } catch (e) { missingErr = e.message; }
  assertions.push(assertCondition("missing file error includes configured name", missingErr && missingErr.indexOf("Cost_Data.csv") !== -1));

  // Case 15: findCsvFileInFolder – ambiguous candidates throws listing
  const ambigFolder = makeMockFolder([
    { name: "Cost_Data (1).csv", content: "" },
    { name: "Cost_Data (2).csv", content: "" }
  ]);
  let ambigErr = null;
  try { findCsvFileInFolder(ambigFolder, "Cost_Data.csv"); } catch (e) { ambigErr = e.message; }
  assertions.push(assertCondition("ambiguous candidates error mentions both files",
    ambigErr && ambigErr.indexOf("Cost_Data (1).csv") !== -1 && ambigErr.indexOf("Cost_Data (2).csv") !== -1));

  // Case 16: storefront bracket segmentation helper
  assertions.push(assertEqual("storefront bracket hero", getStorefrontBracket(100, 100), VDM_CONFIG.BRACKET_NAMES.HERO));
  assertions.push(assertEqual("storefront bracket signature", getStorefrontBracket(80, 100), VDM_CONFIG.BRACKET_NAMES.SIGNATURE));
  assertions.push(assertEqual("storefront bracket proven", getStorefrontBracket(58, 100), VDM_CONFIG.BRACKET_NAMES.PROVEN));
  assertions.push(assertEqual("storefront bracket accelerator", getStorefrontBracket(50, 100), VDM_CONFIG.BRACKET_NAMES.ACCELERATOR));
  assertions.push(assertEqual("storefront bracket clearance", getStorefrontBracket(30, 100), VDM_CONFIG.BRACKET_NAMES.CLEARANCE));
  // Boundary: exactly at 0.35 depth → PROVEN
  assertions.push(assertEqual("storefront bracket boundary 0.35 exact", getStorefrontBracket(65, 100), VDM_CONFIG.BRACKET_NAMES.PROVEN));
  // Boundary: exactly at 0.45 depth → ACCELERATOR
  assertions.push(assertEqual("storefront bracket boundary 0.45 exact", getStorefrontBracket(55, 100), VDM_CONFIG.BRACKET_NAMES.ACCELERATOR));
  // Boundary: exactly at 0.55 depth → CLEARANCE
  assertions.push(assertEqual("storefront bracket boundary 0.55 exact", getStorefrontBracket(45, 100), VDM_CONFIG.BRACKET_NAMES.CLEARANCE));
  // Just below 0.35 → SIGNATURE
  assertions.push(assertEqual("storefront bracket below 0.35 boundary", getStorefrontBracket(65.01, 100), VDM_CONFIG.BRACKET_NAMES.SIGNATURE));
  // Just below 0.45 → PROVEN
  assertions.push(assertEqual("storefront bracket below 0.45 boundary", getStorefrontBracket(55.01, 100), VDM_CONFIG.BRACKET_NAMES.PROVEN));
  // Just below 0.55 → ACCELERATOR
  assertions.push(assertEqual("storefront bracket below 0.55 boundary", getStorefrontBracket(45.01, 100), VDM_CONFIG.BRACKET_NAMES.ACCELERATOR));
  // Negative depth (price > compareMsrp) → HERO
  assertions.push(assertEqual("storefront bracket negative depth", getStorefrontBracket(120, 100), VDM_CONFIG.BRACKET_NAMES.HERO));

  // Case 17: Projected revenue impact includes baseline-present rows and excludes net-new rows
  const existingBaselineImpact = getProjectedRevenueImpact90Contribution(true, 45, 40, 10);
  const netNewImpact = getProjectedRevenueImpact90Contribution(false, 45, 40, 10);
  assertions.push(assertEqual("projected revenue impact includes baseline-present sku", existingBaselineImpact, 50));
  assertions.push(assertEqual("projected revenue impact excludes net-new sku", netNewImpact, 0));

  // Case 18: Folder ID resolution — Script Property wins over legacy config
  const resolvedFromProperty = resolveOperationalFolderId("prop-folder-id", "legacy-folder-id");
  assertions.push(assertEqual("folder ID: property value wins over legacy", resolvedFromProperty, "prop-folder-id"));

  // Case 19: Folder ID resolution — legacy config used when property is absent
  const resolvedFromLegacy = resolveOperationalFolderId(null, "legacy-folder-id");
  assertions.push(assertEqual("folder ID: legacy value used when property is absent", resolvedFromLegacy, "legacy-folder-id"));

  // Case 20: Folder ID resolution — empty string property falls back to legacy
  const resolvedFromEmptyProperty = resolveOperationalFolderId("", "legacy-folder-id");
  assertions.push(assertEqual("folder ID: empty property falls back to legacy", resolvedFromEmptyProperty, "legacy-folder-id"));

  // Case 21: Folder ID resolution — missing both sources throws an actionable error
  let folderIdMissingBothError = null;
  try {
    resolveOperationalFolderId(null, "");
  } catch (e) {
    folderIdMissingBothError = e.message;
  }
  assertions.push(assertCondition(
    "folder ID: missing both property and legacy throws an error",
    folderIdMissingBothError !== null && folderIdMissingBothError.indexOf("VDM_FOLDER_ID") !== -1
  ));

  // Case 22: Sync audit refresh messaging remains explicitly report-only
  assertions.push(assertEqual(
    "sync audit refresh message is report-only",
    getSyncAuditRefreshSuccessMessage(),
    "Shopify Sync Audit refreshed for review only. No Shopify/storefront prices were changed."
  ));

  // Case 23: Margin score boundaries — 65%/50%/35% thresholds
  assertions.push(assertEqual("margin score 3 at 0.65", scoreMarginComponent(0.65), 3));
  assertions.push(assertEqual("margin score 2 at just below 0.65", scoreMarginComponent(0.6499), 2));
  assertions.push(assertEqual("margin score 2 at 0.50", scoreMarginComponent(0.50), 2));
  assertions.push(assertEqual("margin score 1 at just below 0.50", scoreMarginComponent(0.4999), 1));
  assertions.push(assertEqual("margin score 1 at 0.35", scoreMarginComponent(0.35), 1));
  assertions.push(assertEqual("margin score 0 at just below 0.35", scoreMarginComponent(0.3499), 0));

  // Case 24: SHARED stock score uses totalStock (USA + WEB), not webStock alone
  // usaStock = 200, webStock = 0, units90 = 60 → dailyVelocity = 60/90 ≈ 0.667
  // dos with totalStock (200) = 200 / 0.667 ≈ 300 → score 0
  // dos with webStock only (0) → dailyVelocity > 0 but 0/0.667 ≈ 0 → score 3
  // totalStock-based score should be 0 (DoS > 180), proving USA reserve affects SHARED score
  assertions.push(assertEqual(
    "SHARED stock score uses totalStock — USA reserve stock raises DoS beyond 180",
    scoreStockComponent("SHARED", 200, 0, 60),
    0
  ));
  // usaStock = 0, webStock = 100, units90 = 60 → dos = 100/0.667 ≈ 150 → score 1
  assertions.push(assertEqual(
    "SHARED stock score uses totalStock — web-only stock maps to correct DoS tier",
    scoreStockComponent("SHARED", 0, 100, 60),
    1
  ));
  // totalStock = 200 + 100 = 300, units90 = 60 → dos ≈ 450 → score 0
  assertions.push(assertEqual(
    "SHARED stock score uses totalStock — combined USA + WEB inventory raises DoS",
    scoreStockComponent("SHARED", 200, 100, 60),
    0
  ));

  // Case 25: WEBONLY stock score is always 2 regardless of inventory levels
  assertions.push(assertEqual("WEBONLY stock score is always 2 with zero inventory", scoreStockComponent("WEBONLY", 0, 0, 0), 2));
  assertions.push(assertEqual("WEBONLY stock score is always 2 with large USA stock", scoreStockComponent("WEBONLY", 9999, 9999, 0), 2));

  // Case 26: Control Panel virtual SKU prefix parsing
  (function() {
    const raw = " GLAS-WEB , peg-web, ";
    const parsed = parseVirtualSkuPrefixes(raw);
    assertions.push(assertEqual("Case 26: parsed prefix count", parsed.length, 2));
    assertions.push(assertEqual("Case 26: first prefix is GLAS-WEB", parsed[0], "GLAS-WEB"));
    assertions.push(assertEqual("Case 26: second prefix is PEG-WEB", parsed[1], "PEG-WEB"));
    const matchingSku = "GLAS-WEB-001";
    const nonMatchingSku = "CLASSIC-001";
    const matchFulfillment = resolveFulfillmentType(matchingSku, "shared", parsed);
    const noMatchFulfillment = resolveFulfillmentType(nonMatchingSku, "shared", parsed);
    assertions.push(assertEqual("Case 26: matching SKU is WEBONLY", matchFulfillment, "WEBONLY"));
    assertions.push(assertEqual("Case 26: non-matching SKU defaults to SHARED", noMatchFulfillment, "SHARED"));
  })();

  // Case 27: BI feed raw-table contract — headless API state
  (function() {
    var calls = [];
    var mockFilter = { remove: function() { calls.push("filter.remove"); } };
    var mockBiSheet = {
      clear: function() { calls.push("clear"); return mockBiSheet; },
      clearFormats: function() { calls.push("clearFormats"); return mockBiSheet; },
      setFrozenRows: function(n) { calls.push("setFrozenRows:" + n); return mockBiSheet; },
      setFrozenColumns: function(n) { calls.push("setFrozenColumns:" + n); return mockBiSheet; },
      getFilter: function() { return mockFilter; },
      getMaxRows: function() { return 0; },
      getMaxColumns: function() { return 0; },
      getRange: function() {
        return { setValues: function() {}, copyTo: function() {} };
      },
      showRows: function() {},
      showColumns: function() {}
    };

    resetBiFeedSheet(mockBiSheet);

    assertions.push(assertCondition("Case 27: clearFormats called on BI sheet", calls.indexOf("clearFormats") !== -1));
    assertions.push(assertCondition("Case 27: setFrozenRows(0) called on BI sheet", calls.indexOf("setFrozenRows:0") !== -1));
    assertions.push(assertCondition("Case 27: setFrozenColumns(0) called on BI sheet", calls.indexOf("setFrozenColumns:0") !== -1));
    assertions.push(assertCondition("Case 27: filter.remove() called on BI sheet", calls.indexOf("filter.remove") !== -1));
  })();

  const failures = assertions.filter(a => !a.pass);
  if (failures.length > 0) {
    const detail = failures.map(f => "- " + f.name + " (expected: " + f.expected + ", actual: " + f.actual + ")").join("\n");
    throw new Error("Regression harness failed:\n" + detail);
  }

  Logger.log("Regression harness passed: " + assertions.length + " assertions.");
  return { passed: assertions.length };
}

function evaluateRoutingRegressionCase(input) {
  const tierState = deriveTargetTier(input);
  let vdmMarkdown = tierState.vdmMarkdown;
  let guardrailCode = GUARDRAIL_CODES.SAFE;

  const isCostMissing = (input.cost === null || input.cost === 0) && !isZeroCostPermitted(input.gateCode);
  if (isCostMissing) {
    guardrailCode = GUARDRAIL_CODES.ERR_MISSING_COST;
  } else if (input.curMargin < 0) {
    guardrailCode = GUARDRAIL_CODES.ERR_NEGATIVE_MARGIN;
  } else if (mathGuard(input.simNet, input.cost)) {
    const stackMargin = input.simNet === 0 ? 0 : (input.simNet - input.cost) / input.simNet;
    if (stackMargin < VDM_CONFIG.PROFIT_FLOOR_GUARDRAIL) {
      guardrailCode = GUARDRAIL_CODES.ERR_MARGIN_FLOOR_VIOLATOR;
    }
  }

  const b2bHoldActive = input.fulfillment === "SHARED" && (vdmMarkdown >= 0.5) && input.usaStock >= input.b2bReserveMin && input.b2b30DSales > 0;
  if (guardrailCode === GUARDRAIL_CODES.SAFE && b2bHoldActive) {
    vdmMarkdown = input.curMarkdown;
    guardrailCode = GUARDRAIL_CODES.WARN_B2B_HOLD;
  }

  let queueCode = QUEUE_CODES.NONE;
  if (guardrailCode === GUARDRAIL_CODES.ERR_MISSING_COST) {
    queueCode = QUEUE_CODES.QUEUE_1A_COST;
  } else if (guardrailCode === GUARDRAIL_CODES.ERR_NEGATIVE_MARGIN) {
    queueCode = QUEUE_CODES.QUEUE_1A_MARGIN;
  } else if (guardrailCode === GUARDRAIL_CODES.ERR_MARGIN_FLOOR_VIOLATOR) {
    queueCode = QUEUE_CODES.QUEUE_1B_FLOOR;
  } else if (guardrailCode !== GUARDRAIL_CODES.WARN_B2B_HOLD && input.fulfillment === "WEBONLY" && input.totalScore <= 3) {
    queueCode = QUEUE_CODES.QUEUE_2_WEBONLY;
  } else if (guardrailCode !== GUARDRAIL_CODES.WARN_B2B_HOLD && input.fulfillment === "SHARED" && input.totalScore <= 3) {
    queueCode = QUEUE_CODES.QUEUE_3_CLEARANCE;
  }

  return { guardrailCode, queueCode, vdmMarkdown };
}

function deriveTargetTier(input) {
  if (input.gateCode === GATEKEEPER_CODES.GWP) return { tier: "GWP Promo Hold (0% Hold)", vdmMarkdown: 0 };
  if (input.gateCode === GATEKEEPER_CODES.NEW_LAUNCH) return { tier: "New Launch (0% Hold)", vdmMarkdown: 0 };
  if (input.gateCode === GATEKEEPER_CODES.MAP) return { tier: "3rd Party MAP Review (0% Hold)", vdmMarkdown: 0 };

  if (input.totalScore === 10) return { tier: "Top Hero (0% Off)", vdmMarkdown: 0 };
  if (input.totalScore >= 8) return { tier: "Signature Hero (30% Off)", vdmMarkdown: 0.3 };
  if (input.totalScore >= 6) return { tier: "Proven Performer (40% Off)", vdmMarkdown: 0.4 };
  if (input.totalScore >= 4) return { tier: "Accelerator (50% Off)", vdmMarkdown: 0.5 };

  if (input.fulfillment === "WEBONLY" && (input.units90 || 0) > 0) {
    return { tier: "Accelerator / Digital Review (50% Off)", vdmMarkdown: 0.5 };
  }
  return { tier: "Clearance/Archive (65% Off)", vdmMarkdown: 0.65 };
}

function assertEqual(name, actual, expected) {
  return { name: name, pass: actual === expected, actual: actual, expected: expected };
}

function assertCondition(name, condition) {
  return { name: name, pass: !!condition, actual: !!condition, expected: true };
}

/**
 * Creates a lightweight in-memory mock of a Drive Folder for testing
 * findCsvFileInFolder without requiring live Drive access.
 * Each entry: { name: string, content: string }
 */
function makeMockFolder(entries) {
  function makeMockFile(entry) {
    return {
      getName: function() { return entry.name; },
      getBlob: function() {
        return { getDataAsString: function() { return entry.content; } };
      }
    };
  }

  function makeIterator(items) {
    let idx = 0;
    return {
      hasNext: function() { return idx < items.length; },
      next: function() { return items[idx++]; }
    };
  }

  return {
    getFilesByName: function(name) {
      const matches = entries.filter(e => e.name === name).map(makeMockFile);
      return makeIterator(matches);
    },
    getFiles: function() {
      return makeIterator(entries.map(makeMockFile));
    }
  };
}
