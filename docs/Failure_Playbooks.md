# Failure Playbooks

## 1) Missing required headers

**Common error**

- `REQUIRED HEADER MISSING: expected one of ...`

**Likely causes**

- Source export schema changed
- Header typo, case drift, or renamed field
- Wrong file dropped into Drive folder

**Resolution**

1. Run **Advanced Diagnostics → Run Pre-Flight Sanity Check**.
2. Compare source headers with accepted aliases in `/home/runner/work/glas-scoring-script/glas-scoring-script/README.md` and `/home/runner/work/glas-scoring-script/glas-scoring-script/Config.gs`.
3. Correct export mapping at source or replace file with correct export.
4. Re-run pre-flight, then full sync.

## 2) Missing sales tab / empty sales tab

**Common error**

- `REQUIRED TAB MISSING: Tab 'shopify_90day_sales' is missing or empty...`

**Likely causes**

- `Total sales by product.csv` not uploaded
- Sales ingestion not run
- Tab was manually modified/cleared

**Resolution**

1. Ensure `Total sales by product.csv` exists in configured Drive folder.
2. Run full sync (or targeted metadata ingest path that includes sales sync).
3. Confirm `shopify_90day_sales` contains headers and data rows.

## 3) EEI warehouse row-5 format failures

**Common errors**

- `contains insufficient rows... Expected metadata rows 1–4, header on row 5, and data on row 6+.`
- Required header missing in EEI files

**Likely causes**

- Metadata/header row positions shifted
- CSV trimmed to fewer than six rows
- Wrong report format exported

**Resolution**

1. Verify rows 1–4 are metadata.
2. Verify row 5 contains `ITEM CODE` and stock header.
3. Verify data starts row 6+.
4. Re-export report in expected format and replace file.
5. Re-run pre-flight.

## 4) Queue 1A missing-cost blocks

**Common guardrail/code**

- `❌ BLOCKED (Missing Cost)`
- `ERR_MISSING_COST`
- Queue `Q1A_MISSING_COST`

**Likely causes**

- Cost waterfall fields blank for SKU
- SKU mismatch between catalog and cost files
- Source cost values malformed/non-numeric

**Resolution**

1. Validate SKU normalization between `shopify_export_gt.csv` and `Cost_Data.csv`.
2. Verify waterfall fields are populated in this precedence:
   - `EEI LAST PURCHASE PRICE` → `GLAS COSTING` → `COTR LAST PURCHASE PRICE` → `COST` → `UNIT COST` → Shopify `COST PER ITEM`.
3. Correct source data and rerun sync.

## 5) Queue 1A negative margin audits

**Common guardrail/code**

- `❌ BLOCKED (Negative Base Margin)`
- `ERR_NEGATIVE_MARGIN`

**Likely causes**

- Live storefront price below resolved cost
- Cost source spike or stale storefront pricing

**Resolution**

1. Compare `Live Storefront Price` and `Resolved Cost Base` in dashboard matrix.
2. Correct cost source or storefront base price.
3. Re-run full sync.

## 6) Queue 1B margin floor violators

**Common guardrail/code**

- `❌ BLOCKED (Margin Floor Violator)`
- `ERR_MARGIN_FLOOR_VIOLATOR`

**Likely causes**

- Proposed markdown + affiliate discount drops stacked margin below 20%

**Resolution**

1. Confirm affiliate rate in `[01] Control Panel` (`E2`).
2. Verify resolved cost and proposed markdown tier.
3. Reduce discount or remediate cost data before execution.

## 7) B2B hold overrides

**Common guardrail/status**

- `⚠️ B2B HOLD`
- `⚠️ HOLD: B2B Volume Stable`

**Expected behavior**

- For qualifying SHARED SKUs, markdown reverts to current live markdown and queue 2/3 is suppressed.

**Resolution steps if unexpected**

1. Check SHARED fulfillment tag.
2. Confirm proposed markdown is at least 50%.
3. Confirm USA stock and USA `SALES PAST 30 DAYS` values.
4. Confirm `B2B Reserve Min Qty` value in control panel column D.
