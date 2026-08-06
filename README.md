# GLAS Scoring Script

Google Apps Script pricing engine for the **Vendor-Driven Markdown (VDM)** workflow.

This script ingests retail, stock, and cost data, calculates a **10-point composite score per SKU**, maps each SKU to a pricing tier, applies guardrails/overrides, and produces structured outputs for review and execution.

---

## What the script does

For each SKU, the script:

1. Loads source data (sales, catalog/pricing, warehouse stock, cost waterfall).
2. Resolves cost and computes current gross margin.
3. Calculates three dimension scores:
   - **Velocity** (0–4)
   - **Margin** (0–3)
   - **Stock** (0–3)
4. Sums dimensions into a **composite score (0–10)**.
5. Maps composite score to a strategic markdown tier.
6. Applies gatekeeper overrides (for example, GWP or new launch freezes).
7. Routes outputs into action/review queues.

---

## Scoring model (VDM 10-point)

### 1) Velocity Score (0–4)
Input: `Total sales by product.csv` (ingested into the `shopify_90day_sales` tab for 90-day units sold)

- **0** = 0 units sold
- **1** = exactly 1 unit sold
- **2** = percentile rank < 55th (among SKUs with >1 unit)
- **3** = percentile rank 55th–79th
- **4** = percentile rank ≥ 80th

Uses tie-aware percentile ranking (`PERCENTRANK.INC` style midpoint handling).

### 2) Margin Score (0–3)
Input: gross margin = `(Live Price − Resolved Cost) / Live Price`

- **0** = margin < 35%
- **1** = 35%–44.9%
- **2** = 45%–54.9%
- **3** = ≥ 55%

### 3) Stock Score (0–3)
Input: web on-hand stock + sales velocity (for SHARED DoS)

- **WEBONLY** → fixed **2**
- **SHARED**:
  - **3** = DoS ≤ 30
  - **2** = DoS 31–120
  - **1** = DoS 121–180
  - **0** = DoS > 180

DoS formula for SHARED:

`DoS = Web Stock / (Units90 / 90)`

If `Units90 = 0`, DoS is set to `999` (zero-sales override).

---

## Composite → Tier mapping

- **10** → Top Hero
- **8–9** → Signature Hero
- **6–7** → Proven Performer
- **4–5** → Accelerator
- **0–3** (WEBONLY, sales90 = 0) → Clearance / Archive

---

## Data inputs

Expected source files and workbook tabs include:

- `Total sales by product.csv` → `shopify_90day_sales` — 90-day unit velocity
- `shopify_export_gt.csv` → `shopify_export` — active SKU catalog + live pricing metadata
- `EEI USA Whse Stock Report.csv` → `eei_usa_whse` — USA on-hand + B2B sales context
- `EEI WEB Whse Stock Report.csv` → `eei_web_whse` — web warehouse stock
- `Cost_Data.csv` → `cost_ledger` — procurement cost waterfall

### Accepted CSV header aliases

- `shopify_export_gt.csv`
  - SKU: `VARIANT SKU` or `SKU`
  - Status: `STATUS`
  - Price: `VARIANT PRICE` or `PRICE`
  - Type: `TYPE` or `PRODUCT TYPE`
- `Cost_Data.csv`
  - SKU: `SKU` or `VARIANT SKU`
  - Cost: `COST`, `UNIT COST`, `EEI LAST PURCHASE PRICE`, `GLAS COSTING`, `COTR LAST PURCHASE PRICE`, or `COST PER ITEM`
- `Total sales by product.csv`
  - SKU: `PRODUCT VARIANT SKU`, `VARIANT SKU`, or `SKU`
  - Quantity: `NET QUANTITY`, `NET ITEMS SOLD`, `QTY`, or `QUANTITY`
- `EEI USA Whse Stock Report.csv` / `EEI WEB Whse Stock Report.csv`
  - Headers are read from **row 5** (rows 1–4 metadata, row 6+ data)
  - SKU: `ITEM CODE` or `SKU`
  - Quantity:
    - USA: `EEI USA WAREHOUSE ON HAND STOCK`, `QTY`, `QUANTITY`, or `ON HAND STOCK`
    - WEB: `EEI WEB WAREHOUSE ON HAND STOCK`, `QTY`, `QUANTITY`, or `ON HAND STOCK`

Resolved-cost precedence is: `EEI LAST PURCHASE PRICE` → `GLAS COSTING` → `COTR LAST PURCHASE PRICE` → `COST` → `UNIT COST` → Shopify `COST PER ITEM` → `0`.

---

## Operational guardrails

The script supports gatekeeper holds that override standard tier markdown behavior, including:

- Active GWP promo hold (markdown freeze)
- New launch hold (markdown freeze)

When a gatekeeper condition is met, pricing action is held regardless of composite score.

---

## Running / maintenance notes

- Built for Google Apps Script operational workflows.
- Keep workbook tab names stable (especially `shopify_90day_sales` for `Total sales by product.csv`) to avoid broken mappings.
- Treat this README as the behavior contract; update it when scoring logic or thresholds change.

---

## Related documentation

- `docs/VDM_Master_Specification.md`
- `docs/VDM_Scoring_Framework.md`

These docs define the production scoring framework and should stay in sync with script behavior.
