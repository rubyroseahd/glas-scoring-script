# VDM Scoring Framework
**Version:** 3.0.0 — Production Locked  
**System:** Vendor-Driven Markdown (VDM) Pricing Engine

---

## Scoring Model Summary

The VDM Engine uses a **10-point composite scoring model** combining three independently scored dimensions: Velocity (0–4), Margin (0–3), and Stock (0–3). The composite score drives tier assignment and discount depth.

---

## Dimension 1: Retail Velocity Score (0–4)

**Input:** `Total sales by product.csv` — 90-day net units sold per SKU.

**Method:** Tie-aware percentile ranking (`PERCENTRANK.INC` style). Percentile is computed across all SKUs with more than 1 unit sold, using interpolated midpoint ranking within tie groups to prevent under-ranking of duplicates.

| Score | Bracket | Condition |
|---|---|---|
| **0** | No Velocity | 0 units sold (90-day total) |
| **1** | Minimal Velocity | Exactly 1 unit sold |
| **2** | Low Velocity | Percentile rank < 55th percentile |
| **3** | Mid Velocity | Percentile rank 55th–79th percentile |
| **4** | High Velocity | Percentile rank ≥ 80th percentile |

> SKUs with 0 sales are excluded from the percentile distribution; they receive Score 0 directly.

---

## Dimension 2: Margin Score (0–3)

**Input:** Current gross margin = `(Live Price − Resolved Cost) / Live Price`

| Score | Bracket | Condition |
|---|---|---|
| **0** | Below Floor | Gross margin < 35% |
| **1** | Acceptable | Gross margin 35%–44.9% |
| **2** | Healthy | Gross margin 45%–54.9% |
| **3** | Premium | Gross margin ≥ 55% |

---

## Dimension 3: Retail Stock Score (0–3)

**Input:** EEI Web Warehouse on-hand stock + 90-day sales velocity for DoS calculation.

**DoS Formula (SHARED only):**  
`DoS = Web Stock / (Units90 / 90)`  
If `Units90 = 0`, DoS is set to **999** (zero-sales override — no stock pressure signal).

| Fulfillment | Score | Bracket | Condition |
|---|---|---|---|
| WEBONLY | **2** | Virtual Stock | Always (no physical DoS pressure) |
| SHARED | **3** | Critical Low | DoS ≤ 30 days |
| SHARED | **2** | Healthy | DoS 31–120 days |
| SHARED | **1** | Elevated | DoS 121–180 days |
| SHARED | **0** | Overstocked | DoS > 180 days |

---

## Composite Score & Strategic Tier Mapping

| Total Score | Tier | Base VDM Markdown | Affiliate Stacked Total* |
|---|---|---|---|
| 10 | Top Hero | 0% | ~15% |
| 8–9 | Signature Hero | 30% | ~40.5% |
| 6–7 | Proven Performer | 40% | ~49% |
| 4–5 | Accelerator | 50% | ~57.5% |
| 0–3 (WEBONLY, sales90 = 0) | Clearance / Archive | 65% | ~70.25% |
| 0–3 (WEBONLY, sales90 > 0) | Accelerator / Digital Review | 50% | ~57.5% |
| 0–3 (SHARED) | Clearance / Archive | 65% | ~70.25% |

*Stacked total = `1 − (1 − Base Markdown) × (1 − Affiliate Rate)` at default 15% affiliate rate.

---

## WEBONLY Low-Score Exception (Score 0–3)

When a WEBONLY SKU scores 0–3, the system applies a two-path decision:

> **90-day total sales = 0** is used as the operational proxy for three consecutive 30-day zero-sales periods, since the `Total sales by product.csv` source provides a single 90-day aggregate figure.

| Condition | Path | Tier | Markdown |
|---|---|---|---|
| `sales90 = 0` | Clearance-eligible | Clearance / Archive | 65% off |
| `sales90 > 0` | Hold for review | Accelerator / Digital Review | 50% off |

This prevents premature clearance of items that still show some sales momentum within the period.

---

## Gatekeeper Overrides (Pre-Scoring)

Gatekeepers intercept the tier assignment before composite scoring applies.

| Gatekeeper | Trigger | Result |
|---|---|---|
| ⚠️ Active GWP Promo | SKU in GWP registry (column A of Settings) | Tier: GWP Promo Hold; Markdown: 0% |
| New Launch | SKU in New Launch registry (column B of Settings) | Tier: New Launch Hold; Markdown: 0% |
| 3rd Party MAP | Vendor string exactly matches MAP registry entry (column C of Settings) | Tier: MAP Review Hold; Markdown: 0% |

**MAP matching is vendor-level only:** The vendor string must exactly equal a registered entry in the MAP Restricted Brands list. Partial substring matches are not used.

---

## Profit Guardrail Logic

### Stacked Margin Calculation

```
Proposed Price = Compare MSRP × (1 − VDM Markdown %)
Simulated Net  = Proposed Price × (1 − Affiliate Rate)
Stacked Margin = (Simulated Net − Resolved Cost) / Simulated Net
```

### Guardrail Thresholds

| Status | Condition | Queue |
|---|---|---|
| ✓ SAFE | Stacked Margin ≥ 20% | No action |
| ❌ BLOCKED | Stacked Margin < 20% | Queue 1B |
| DATA_ERROR | Cost data missing | Manual review required |

---

## Action Queue Priority

Queues are evaluated in priority order — a SKU can only be in one queue.

| Priority | Queue | Trigger | Action |
|---|---|---|---|
| 1 | **Queue 1A** | Current gross margin < 0% | ❌ BLOCKED — Negative base margin audit |
| 2 | **Queue 1B** | Simulated stacked margin < 20% | ❌ BLOCKED — Checkout margin guardrail |
| 3 | **Queue 2** | WEBONLY + Score 0–3 | WEBONLY digital review |
| 4 | **Queue 3** | SHARED + Score 0–3 | SHARED clearance / liquidation |

---

## Compound Stacking Example

**Scenario:** SKU with MSRP $100, Proven Performer tier (40% off), 15% affiliate rate.

```
Proposed Price     = $100 × (1 − 0.40)  = $60.00
Simulated Net      = $60 × (1 − 0.15)   = $51.00
Stacked Discount   = 1 − (0.60 × 0.85)  = 49%
```

If resolved cost = $40:  
```
Stacked Margin = ($51 − $40) / $51 = 21.6%  →  ✓ SAFE
```

If resolved cost = $44:  
```
Stacked Margin = ($51 − $44) / $51 = 13.7%  →  ❌ BLOCKED (Queue 1B)
```
