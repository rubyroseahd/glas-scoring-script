# VDM Master Specification
**Version:** 3.0.0 — Production Locked  
**System:** Vendor-Driven Markdown (VDM) Pricing Engine

---

## 1. Overview

The VDM Engine is a Google Apps Script–powered pricing operations system that ingests multi-source retail data, computes a 10-point composite score per SKU, assigns strategic discount tiers, enforces profit guardrails, and routes action items to structured queues for operational review.

---

## 2. Data Sources

| Source File | Tab Key | Purpose |
|---|---|---|
| `Total sales by product.csv` | `_raw_sales` | 90-day retail velocity (units sold) |
| `shopify_export_gt.csv` | `_raw_shopify` | Active SKU catalog, pricing, vendor metadata |
| `EEI USA Whse Stock Report.csv` | `_raw_eei_usa` | USA warehouse on-hand stock + 30-day B2B sales |
| `EEI WEB Whse Stock Report.csv` | `_raw_eei_web` | Web warehouse on-hand stock |
| `Cost_Data.csv` | `_raw_cost` | Procurement cost waterfall (EEI → Gläs → COTR → Shopify) |

> **Note:** The sales input tab uses the file `Total sales by product.csv` consistently throughout the system.

---

## 3. 10-Point VDM Scoring Engine

### 3.1 Velocity Score (0–4)

Measures 90-day retail unit velocity using **tie-aware percentile ranking** (`PERCENTRANK.INC` style).

| Score | Condition |
|---|---|
| 0 | 0 units sold |
| 1 | Exactly 1 unit sold |
| 2 | Percentile rank < 55th (among SKUs with >1 unit) |
| 3 | Percentile rank 55th–79th |
| 4 | Percentile rank ≥ 80th |

**Tie handling:** Uses interpolated INC-style midpoint ranking within tie groups, ensuring duplicate sales values are scored fairly rather than being under-ranked.

### 3.2 Margin Score (0–3)

Measures current gross margin based on live Shopify price vs. resolved procurement cost.

| Score | Condition |
|---|---|
| 0 | Margin < 35% |
| 1 | Margin 35%–44.9% |
| 2 | Margin 45%–54.9% |
| 3 | Margin ≥ 55% |

### 3.3 Stock Score (0–3)

Measures inventory pressure by fulfillment channel.

| Fulfillment | Score | Condition |
|---|---|---|
| WEBONLY | 2 | Always (no physical stock pressure) |
| SHARED | 3 | Days of Supply (DoS) ≤ 30 |
| SHARED | 2 | DoS 31–120 |
| SHARED | 1 | DoS 121–180 |
| SHARED | 0 | DoS > 180 |

**Zero-sales DoS override:** If daily velocity = 0, DoS is set to 999 (no stock pressure signal).

---

## 4. Gatekeeper Overrides

Gatekeepers bypass scoring tiers and apply fixed pricing holds.

| Gatekeeper | Trigger | Behavior |
|---|---|---|
| ⚠️ Active GWP Promo | SKU in GWP registry | Pricing frozen; markdown held at 0% |
| New Launch | SKU in New Launch registry | Pricing frozen; markdown held at 0% |
| 3rd Party MAP | **Vendor string** matches MAP registry (exact vendor-level match) | Pricing frozen; MAP Review hold at 0% |

> **MAP Enforcement:** MAP matching is enforced at the **vendor level only**. The vendor string from Shopify is matched exactly against entries in the MAP Restricted Brands column of the Settings tab (column C). SKU-level MAP entries are not used.

---

## 5. Strategic Tier Mapping

| Total Score | Tier | VDM Markdown |
|---|---|---|
| 10 | Top Hero | 0% |
| 8–9 | Signature Hero | 30% |
| 6–7 | Proven Performer | 40% |
| 4–5 | Accelerator | 50% |
| 0–3 (WEBONLY, sales90 = 0) | Clearance / Archive | 65% |
| 0–3 (WEBONLY, sales90 > 0) | Accelerator / Digital Review | 50% |
| 0–3 (SHARED) | Clearance / Archive | 65% |

### WEBONLY Exception (Score 0–3)

When a WEBONLY SKU scores 0–3:

- If **90-day total sales = 0** (used as the operational proxy for three consecutive 30-day zero-sales periods), the SKU is eligible for Clearance / Archive at 65% off.
- If **90-day total sales > 0**, the SKU is held at Accelerator / Digital Review at 50% off for further review before committing to clearance.

> **Operational proxy note:** The system uses 90-day total sales = 0 as the proxy for three consecutive 30-day zero-sales periods, since the source data provides a single 90-day aggregate figure from `Total sales by product.csv`.

---

## 6. Profit Guardrails

### 6.1 Queue 1B — Simulated Checkout Margin Guardrail

Before any markdown is applied to the storefront, the engine simulates the checkout margin:

```
Proposed Price = Compare MSRP × (1 − VDM Markdown %)
Simulated Net  = Proposed Price × (1 − Affiliate Rate)
Stacked Margin = (Simulated Net − Cost) / Simulated Net
```

If `Stacked Margin < 20%`, the SKU is flagged **❌ BLOCKED** and routed to **Queue 1B**.

### 6.2 B2B Volume Stable Override

If a SHARED SKU has:
- VDM Markdown ≥ 50%
- USA Warehouse stock ≥ **B2B Reserve Min Qty** (Settings column D; defaults to 500 if not configured)
- 30-day B2B sales > 0

...then the pricing migration is held at the current live price (`⚠️ HOLD: B2B Volume Stable`) and the tier is changed to **B2B Protection Hold**.

---

## 7. Action Queue Structure

| Queue | Trigger | Action Required |
|---|---|---|
| **Queue 1A** | Current gross margin < 0 (negative base margin) | ❌ BLOCKED — Flag for cost audit; pricing cannot proceed |
| **Queue 1B** | Simulated stacked checkout margin < 20% | ❌ BLOCKED — Pricing blocked by profit guardrail |
| **Queue 2** | WEBONLY + Total Score 0–3 | Route to WEBONLY digital review; evaluate clearance vs. hold |
| **Queue 3** | SHARED + Total Score 0–3 | Route to SHARED clearance/liquidation channel |

Queue 1A is evaluated first (takes precedence), followed by Queue 1B, then Queue 2/3.

---

## 8. Dashboard Matrix Columns (28 Columns)

| # | Column Header | Description |
|---|---|---|
| A | SKU Anchor Key | Normalized SKU (uppercase) |
| B | Gatekeeper Status | GWP / New Launch / MAP / None |
| C | Fulfillment Tag | WEBONLY or SHARED |
| D | Resolved Cost Base | Procurement cost after waterfall |
| E | Live Storefront Price | Current Shopify price |
| F | Live Compare MSRP | Reference MSRP for markdown math |
| G | Active Storefront Markdown Depth % | Current live markdown depth |
| H | Current Gross Margin % | (Price − Cost) / Price |
| I | Raw 90D Retail Velocity | Units sold in 90 days |
| J | Retail Velocity Score Component | 0–4 |
| K | Margin Score Component | 0–3 |
| L | Retail Stock Score Component | 0–3 |
| M | Total Composite Score | Sum of J + K + L |
| N | Target Strategic Tier | VDM tier label |
| O | VDM Markdown Depth % | Target markdown depth |
| P | Total On-Hand Warehouse Stock | USA + Web warehouse total |
| Q | EEI Web Warehouse On Hand Stock | Web warehouse units |
| R | Live Storefront Shopify Qty | Shopify inventory count |
| S | Asynchronous Inventory Drift Tracker | Shopify Qty − Web Stock |
| T | New Proposed Storefront Price | MSRP × (1 − VDM Markdown) |
| U | Simulated Checkout Net Price | Proposed Price × (1 − Affiliate Rate) |
| V | Final Simulated Stacked Margin % | (Net − Cost) / Net |
| W | Profit Guardrail Status Alert | ✓ SAFE / ❌ BLOCKED / DATA_ERROR |
| X | Current Equivalent Storefront Tier | Current markdown tier label |
| Y | Pricing Migration Status | Price Hold / Deepen Discount / Price Recovery |
| Z | Retail Price Shift ($) | Proposed Price − Current Price |
| AA | Net Margin Change % | Stacked Margin − Current Margin |
| AB | Action Queue | Queue 1A / 1B / 2 / 3 routing label |

---

## 9. Settings Tab (Column Reference)

| Column | Registry |
|---|---|
| A | Active GWP SKUs |
| B | New Launch Override SKUs |
| C | MAP Restricted Brands (vendor-level strings) |
| D | B2B Reserve Minimum Quantity |
| E | Affiliate Coupon Rate (default: 15%) |

---

## 10. Monthly SOP

1. Upload source files to the configured Google Drive folder.
2. Open the spreadsheet and click **EEI Pricing Engine Launcher → 1. Full System Sync (Standard)**.
3. Review the **[00] Action Items & Sign-off** tab for Queue 1A/1B/2/3 items.
4. Approve or escalate flagged SKUs before pushing changes to Shopify.
5. Use **Advanced Diagnostics → Run Pre-Flight Sanity Check** to validate source file headers before any sync.
