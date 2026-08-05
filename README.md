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
Input: `Total sales by product.csv` (ingested into the `_raw_sales` tab for 90-day units sold)

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

Expected source tabs/files include:

- `Total sales by product.csv` (`_raw_sales`) — 90-day unit velocity
- `shopify_export_gt.csv` (`_raw_shopify`) — active SKU catalog + live pricing metadata
- `EEI USA Whse Stock Report.csv` (`_raw_eei_usa`) — USA on-hand + B2B sales context
- `EEI WEB Whse Stock Report.csv` (`_raw_eei_web`) — web warehouse stock
- `Cost_Data.csv` (`_raw_cost`) — procurement cost waterfall

---

## Operational guardrails

The script supports gatekeeper holds that override standard tier markdown behavior, including:

- Active GWP promo hold (markdown freeze)
- New launch hold (markdown freeze)

When a gatekeeper condition is met, pricing action is held regardless of composite score.

---

## Running / maintenance notes

- Built for Google Apps Script operational workflows.
- Keep source tab names stable (especially `_raw_sales` for `Total sales by product.csv`) to avoid broken mappings.
- Treat this README as the behavior contract; update it when scoring logic or thresholds change.

---

## Related documentation

- `docs/VDM_Master_Specification.md`
- `docs/VDM_Scoring_Framework.md`

These docs define the production scoring framework and should stay in sync with script behavior.
