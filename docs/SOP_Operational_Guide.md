# Standard Operating Procedure (SOP): VDM Pricing Engine Operations

**System:** Vendor-Driven Markdown (VDM) Pricing Engine  
**Version:** 3.0.0-PROD  
**Audience:** Pricing Operations, Ecommerce, Merchandising, and Data Operations

---

## 1. System Overview & Codebase Structure Map

The VDM Pricing Engine is a Google Apps Script–based operational pricing system. It ingests five CSV exports from the configured Google Drive folder, normalizes them into workbook staging tabs, calculates a 10-point SKU score, applies tier and governance logic, enforces profit guardrails, and produces operator-facing action queues and reporting tabs.

### Top-level Google Apps Script modules

| File | Operational responsibility |
|---|---|
| `Config.gs` | Global version, Drive folder ID, source-file names, workbook-tab registry, header registries, defaults, design tokens, and machine-code constants. |
| `Ingestion.gs` | Pre-flight header validation; CSV loading; Shopify, EEI warehouse, 90-day sales, and cost ingestion; normalized staging-tab writes; and the cost-resolution waterfall. |
| `MatrixEngine.gs` | SKU-level scoring, gatekeeper overrides, tier selection, price simulation, margin and B2B guardrails, action-queue assignment, dashboard write, and recovery-point backup. |
| `Utilities.gs` | Shared sheet helpers, safe string/number conversion, MAP-vendor matching, and tie-aware percentile math. |
| `UI.gs` | The `EEI Pricing Engine Launcher` spreadsheet menu, targeted workflows, pre-flight action, sync-only action, rollback, and architecture reset. |
| `AnalyticsReporting.gs` | Creation of the executive summary, action-items/sign-off queue, sync audit, master ledger, supplier scorecard, elasticity ledger, and reporting workflow actions. |

### Documentation suite

| File | Purpose |
|---|---|
| `docs/VDM_Master_Specification.md` | Authoritative operational and business specification for source data, scoring, tiers, guardrails, settings, dashboard outputs, and the monthly operating cycle. |
| `docs/VDM_Scoring_Framework.md` | Detailed scoring taxonomy, ingestion mapping, thresholds, WEBONLY exception, gatekeepers, profit logic, and queue priority. |
| `docs/SOP_Operational_Guide.md` | This operator runbook: staging, execution, queue triage, UAT, and post-go-live backlog. |

### Primary workbook tabs

The system uses the following user-facing tabs: `[00] Executive Brief`, `[00] Action Items & Sign-off`, `[01] Control Panel`, `[01] Supplier Scorecard & Capital Velocity`, `[02] Dashboard Matrix`, `[03] Tier Summary & Distribution Panel`, `[04] Pricing Elasticity & Lift Analytics Ledger`, `[07] Storefront Update & Sync Audit`, and `[09] Master Pricing & Margin Ledger`.

The ingestion/staging layer uses `shopify_export`, `shopify_90day_sales`, `eei_usa_whse`, `eei_web_whse`, and `cost_ledger`. The hidden `_backup_matrix_data` tab is the recovery point used by the emergency rollback workflow.

---

## 2. Data Ingestion & Staging Contract

Place one current copy of each required CSV into the Google Drive folder configured by `VDM_CONFIG.FOLDER_ID`. Source file names are exact and case-sensitive operational identifiers.

### Required source-file-to-workbook-tab mapping

| Source CSV file | Required workbook staging tab | Purpose |
|---|---|---|
| `Total sales by product.csv` | `shopify_90day_sales` | 90-day retail velocity / net units sold. |
| `shopify_export_gt.csv` | `shopify_export` | Active Shopify SKU catalog, live prices, compare-at prices, vendor metadata, inventory, and fulfillment classification source. |
| `Cost_Data.csv` | `cost_ledger` | Cost inputs and resolved-cost staging via the procurement waterfall. |
| `EEI WEB Whse Stock Report.csv` | `eei_web_whse` | Web warehouse on-hand inventory; used for stock-score and days-of-supply calculations. |
| `EEI USA Whse Stock Report.csv` | `eei_usa_whse` | USA warehouse on-hand inventory and 30-day B2B sales signal. |

> **Do not manually reformat normalized staging tabs during a campaign.** The sync overwrites these tabs. Correct the source CSV or Control Panel registry, then rerun the workflow.

### Header and row-offset rules

#### `shopify_export_gt.csv`

- Headers must be on **row 1**.
- Required validation headers:
  - SKU: `VARIANT SKU` or `SKU`
  - Status: `STATUS`
  - Live price: `VARIANT PRICE` or `PRICE`
- The ingestion logic uses additional fields when present, including `HANDLE`, `VARIANT COMPARE AT PRICE`, `VENDOR`, `VARIANT INVENTORY QTY`, `COST PER ITEM`, and `TYPE` or `PRODUCT TYPE`.
- Only unique **active** SKUs are staged. SKU values are normalized to uppercase.
- A product type containing `WEBONLY` becomes `WEBONLY`; otherwise, the fulfillment fallback is `SHARED`.

#### `Total sales by product.csv`

- Headers must be on **row 1**; data starts on **row 2**.
- SKU header accepted by ingestion: `PRODUCT VARIANT SKU`, `VARIANT SKU`, or `SKU`.
- Quantity header accepted by ingestion: `NET QUANTITY`, `NET ITEMS SOLD`, `QTY`, or `QUANTITY`.
- The normalized staging output contains `SKU_ANCHOR`, `Product variant SKU`, and `Net items sold`.
- If the staging tab is missing or empty, the required operator-facing failure wording is:

  `REQUIRED TAB MISSING: Tab 'shopify_90day_sales' is missing or empty. Please ensure 'Total sales by product.csv' has been imported into the workbook.`

#### `Cost_Data.csv`

- Headers must be on **row 1**; data starts on **row 2**.
- Required headers:
  - SKU: `SKU` or `VARIANT SKU`
  - Cost: `COST`, `UNIT COST`, `EEI LAST PURCHASE PRICE`, `GLAS COSTING`, or `COTR LAST PURCHASE PRICE`
- Supported waterfall fields are `EEI LAST PURCHASE PRICE`, `GLAS COSTING`, `COTR LAST PURCHASE PRICE`, `COST`, and `UNIT COST` with Shopify `COST PER ITEM` fallback.
- Resolved cost precedence is: **EEI last purchase price → GLAS costing → COTR last purchase price → Cost → Unit Cost → Shopify Cost per item → 0**.

#### EEI warehouse reports

This rule applies to both `EEI WEB Whse Stock Report.csv` and `EEI USA Whse Stock Report.csv`.

- Rows **1–4** are metadata and are not treated as data.
- The required header row is **row 5**.
- Data must begin on **row 6**.
- Each file must contain at least **6 rows**. Fewer than six rows is a hard ingestion/pre-flight failure.
- Required row-5 headers:
  - SKU: `ITEM CODE` or `SKU`
  - On-hand quantity:
    - USA: `EEI USA WAREHOUSE ON HAND STOCK`, `QTY`, `QUANTITY`, or `ON HAND STOCK`
    - WEB: `EEI WEB WAREHOUSE ON HAND STOCK`, `QTY`, `QUANTITY`, or `ON HAND STOCK`
- `SALES PAST 30 DAYS` is optional in the web report and is used from the USA report for the B2B reserve rule.
- The staging layer normalizes warehouse data to `SKU_ANCHOR`, `ITEM CODE`, on-hand stock, and `SALES PAST 30 DAYS`.

### Pre-flight control

Before a production sync—especially after a source-system export change—run:

**EEI Pricing Engine Launcher → Advanced Diagnostics → Run Pre-Flight Sanity Check**

This validates source availability and the required headers before the engine overwrites staging and reporting tabs.

---

## 3. Monthly Campaign Execution Workflow

### Step 1 — Stage raw CSV exports

1. Export the five required CSVs from Shopify, EEI, and the cost source.
2. Confirm each file name exactly matches the source-file mapping above.
3. Upload or replace the current files in the configured Google Drive folder.
4. Confirm the EEI warehouse files preserve their metadata rows 1–4, headers on row 5, and data from row 6 onward.
5. Run the pre-flight sanity check when file schema, reporting formats, or upstream systems may have changed.

### Step 2 — Audit `[01] Control Panel` settings

Review the `[01] Control Panel` before running pricing. Its row-1 registry headers and row-2/default values govern the campaign:

| Control Panel location | Operational setting | Required use |
|---|---|---|
| `E2` | Affiliate coupon rate | Verify the campaign affiliate rate. The standard default is 15% (`0.15`). This rate is used in simulated checkout net price and stacked-margin calculations. |
| Column A | GWP registry | List active GWP SKUs. Matching SKUs receive `⚠️ Active GWP Promo` and a 0% pricing hold. |
| Column B | New Launch registry | List new-launch override SKUs. Matching SKUs receive a 0% pricing hold. |
| Column C | MAP Restricted Brands | List vendor-level MAP-restricted brands. Review spelling and vendor naming carefully. |
| Column D | B2B Reserve Min Qty | Set the minimum USA warehouse stock quantity for B2B protection. When missing or invalid, the engine defaults to 500. |

MAP is evaluated at the vendor level. The source vendor string is compared against the MAP registry; do not use SKU entries in column C.

### Step 3 — Run the full system sync

Open the workbook and select:

**EEI Pricing Engine Launcher → 1. Run Full System Pricing & Sync**

The standard run performs the following sequence:

1. Validates source schemas.
2. Ingests Shopify, EEI USA, EEI WEB, 90-day sales, and cost data.
3. Resolves costs through the cost waterfall.
4. Recalculates the dashboard matrix.
5. Regenerates the summary, action-items queue, sync audit, master ledger, supplier scorecard, and elasticity snapshot.
6. Writes a backup of the dashboard matrix to `_backup_matrix_data`.

Read the completion alert and record any reported counts for missing costs, negative margins, margin-floor blocks, B2B reserve holds, unmapped physical inventory, and fulfillment fallbacks.

### Step 4 — Triage `[00] Action Items & Sign-off` queues

Open `[00] Action Items & Sign-off` immediately after the run. Review blocked queues before any proposed storefront change is approved for external execution. Use `[02] Dashboard Matrix` for full SKU calculations and `[07] Storefront Update & Sync Audit` to review/export proposed changes versus holds; the sheet itself does not publish anything to Shopify.

---

## 4. Action Queue Triage & Guardrail Precedence

### Strict guardrail precedence

Every SKU has one effective outcome. Apply this hierarchy in strict order:

`ERR_MISSING_COST` > `ERR_NEGATIVE_MARGIN` > `ERR_MARGIN_FLOOR_VIOLATOR` (<20% stacked margin) > `WARN_B2B_HOLD` > `SAFE`

A higher-precedence result suppresses lower-priority routing. In particular, a SKU cannot appear in Queue 2 or Queue 3 when it is already blocked for missing cost, negative margin, or the stacked-margin floor.

### Queue directives

| Queue / code | Trigger | Operator directive |
|---|---|---|
| **Queue 1A — `Q1A_MISSING_COST`** | Cost is missing or zero, unless the SKU is an active GWP item (zero cost is permitted for GWP only). | **Do not price.** Investigate the cost waterfall, SKU normalization, and source cost record. Correct the cost source or approved GWP registry, then rerun the full sync. |
| **Queue 1A — `Q1A_MARGIN`** | Current gross margin is below 0%. | **Do not price.** Flag for a cost and live-price audit. Resolve the incorrect cost or current price before approving any markdown. |
| **Queue 1B — `Q1B_FLOOR`** | Simulated checkout stacked margin is below 20%. | **Do not price at the proposed markdown.** Reduce the discount, revise the applicable affiliate rate if commercially approved, correct cost data, or escalate the exception for approval. |
| **Queue 2 — `Q2_WEBONLY_REVIEW`** | `WEBONLY` SKU with total composite score 0–3, after higher-priority guardrails have cleared. | Conduct digital review. For zero-sales items, approve or reject digital clearance/archive. For items with sales, retain the Accelerator/Digital Review posture pending merchandising review. |
| **Queue 3 — `Q3_SHARED_CLEARANCE`** | `SHARED` SKU with total composite score 0–3, after higher-priority guardrails have cleared. | Route to physical clearance/liquidation planning. Verify inventory, channel implications, and any operational constraints before execution. |

### B2B reserve hold

A `WARN_B2B_HOLD` is not an action queue. It takes precedence over `SAFE` and prevents a deeper markdown when all of the following are true:

- Fulfillment is `SHARED`.
- Candidate VDM markdown is at least 50%.
- USA warehouse stock is at or above the Control Panel B2B reserve minimum (default 500 if not configured).
- USA warehouse `SALES PAST 30 DAYS` is greater than zero.
- No higher-priority cost or margin guardrail applies.

The engine reverts the markdown to the current live markdown, sets the tier to `B2B Protection Hold`, and assigns `⚠️ HOLD: B2B Volume Stable`.

### WEBONLY clearance proxy

For a `WEBONLY` item with a score of 0–3, **`sales90 === 0` is the operational proxy for three consecutive 30-day zero-sales periods before clearance approval.** The sales source is a single 90-day aggregate, so the engine cannot independently prove three monthly zero-sales windows.

- `sales90 === 0` → clearance-eligible `Clearance/Archive (65% Off)`.
- `sales90 > 0` → `Accelerator / Digital Review (50% Off)`; do not treat the SKU as automatically clearance-approved.

---

## 5. 10-Case Operational UAT Verification Suite

Run this suite in a test workbook or against controlled test exports before go-live, after a schema change, and after material engine changes. Record the test input, actual outcome, and pass/fail result.

| # | Scenario | Test setup | Expected result |
|---:|---|---|---|
| 1 | Happy Path | Provide all five valid CSVs, valid Control Panel data, and a normal active SKU with valid cost. | Full sync completes; staging tabs refresh; dashboard and reports generate; the SKU receives a valid score/tier and `SAFE` when its stacked margin is at least 20%. |
| 2 | Missing Master Tab | Remove or clear the cost-resolution/staging dependency before cost resolution. | The run must not silently price from an absent cost basis. Cost resolution reports that required ingestion tabs are missing, or the matrix routes affected SKUs to missing-cost handling rather than approving prices. |
| 3 | Missing Catalog Header | Remove `VARIANT SKU`/`SKU`, `STATUS`, or `VARIANT PRICE`/`PRICE` from `shopify_export_gt.csv`. | Pre-flight or sync fails with a `REQUIRED HEADER MISSING` error naming the accepted header alternatives. |
| 4 | EEI Row Count < 6 | Supply either EEI warehouse CSV with fewer than six rows. | Pre-flight/sync fails and explains that rows 1–4 are metadata, row 5 is the header, and row 6+ is required for data. |
| 5 | EEI Row 5 Header Mismatch | Put warehouse headers on another row or remove `ITEM CODE`/`SKU` or `QTY`/`QUANTITY` from row 5. | Pre-flight/sync fails with the required-header error; headers are not read from rows 1–4 or row 6+. |
| 6 | Sales Schema Mismatch | Remove accepted sales SKU or quantity headers from `Total sales by product.csv`. | Pre-flight/sync fails with a required-header error for accepted SKU and quantity alternatives; no stale sales schema should be accepted. |
| 7 | Zero-Cost Non-GWP | Use a non-GWP SKU with resolved cost `0` or blank. | Guardrail is `ERR_MISSING_COST`; SKU is routed only to Queue 1A `Q1A_MISSING_COST`; no Queue 2/3 duplication occurs. |
| 8 | Zero-Cost GWP Exception | Use a GWP-registry SKU with resolved cost `0`. | Zero cost is permitted for the active GWP exception. The SKU receives the GWP 0% hold rather than an `ERR_MISSING_COST` block, subject to any independently applicable logic. |
| 9 | Queue Precedence & Non-Duplication | Create a low-score WEBONLY or SHARED SKU that also has missing cost, negative current margin, or stacked margin below 20%; separately test a qualifying B2B hold. | Exactly one effective action queue is assigned: missing cost first, then negative margin, then margin-floor violation. B2B hold suppresses Queue 2/3 only when higher guardrails are clear. No SKU appears in multiple queues. |
| 10 | Midpoint Tie Percentile Math | Use multiple SKUs with identical sales values greater than one, positioned around the 55th and 80th percentile thresholds. | Equal sales values receive the same interpolated midpoint percentile from the INC-style tie-aware calculation; tied SKUs must not be arbitrarily under-ranked. |

### UAT evidence checklist

For each passing case, retain:

- The exact source CSV fixture or controlled workbook data.
- Screenshot or export of the relevant staging tab and `[02] Dashboard Matrix` row.
- Relevant `[00] Action Items & Sign-off` row, if queued.
- Completion/error alert text and the operator’s pass/fail sign-off.

---

## 6. Post Go-Live Backlog (v3.1)

The following v3.1 enhancements are planned operational improvements and are not prerequisites for the v3.0.0-PROD workflow:

1. **Preflight schema inspection tool** — provide a detailed source-by-source schema report before ingestion, including detected header row, accepted header match, missing fields, and row-count diagnostics.
2. **Versioned run logger `[99] System Logs`** — create a durable, timestamped log of each run, source-file state, run version, operator, completion status, errors, and summary counters.
3. **Configurable B2B parameters** — expose B2B hold thresholds and related operating parameters through governed configuration rather than fixed assumptions.
4. **Data quality telemetry counters** — surface structured counts for missing cost, unmatched inventory, fulfillment fallbacks, invalid prices, schema exceptions, and other source-data health signals.
5. **Dry-run mode** — calculate staging, score, guardrail, queue, and report outcomes without overwriting production-facing outputs or preparing a storefront change.
6. **Synthetic test suite tab** — add a maintained workbook tab containing reproducible fixtures for the 10-case operational UAT suite and future regression scenarios.

---

## Operator Sign-off

Before approving pricing outputs for a monthly campaign, confirm:

- [ ] All five source CSVs were staged with correct names and schemas.
- [ ] EEI reports retained row-5 headers and row-6 data start.
- [ ] `[01] Control Panel` affiliate rate, GWP, New Launch, MAP, and B2B reserve settings were audited.
- [ ] Full System Sync completed successfully.
- [ ] Queue 1A and Queue 1B items were corrected or formally escalated.
- [ ] Queue 2 and Queue 3 items received merchandising/channel review.
- [ ] B2B holds were reviewed as intentional price-protection outcomes.
- [ ] `[07] Storefront Update & Sync Audit` was reviewed as a report/export artifact before any separate storefront execution.
