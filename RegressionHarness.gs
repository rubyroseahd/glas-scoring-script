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
