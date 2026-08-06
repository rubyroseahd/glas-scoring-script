# Monitoring Metrics Guide

Track these run-health metrics for every full sync.

## Core metrics

1. **Run duration (seconds)**
   - Start timer before `executeFlexibleRefreshProcess`
   - Stop at completion alert
2. **Processed active SKU count**
   - From completion alert: `Processed Active SKUs`
3. **Guardrail pressure metrics**
   - Queue 1A missing cost count
   - Queue 1A negative margin count
   - Queue 1B margin-floor count
   - B2B hold count
4. **Data quality metrics**
   - Unmapped SHARED physical stock count
   - Fulfillment fallback count

## Suggested operating thresholds

- Run duration trend: investigate if run time increases by >25% vs prior 3-run average
- Missing cost: investigate any non-zero count
- Negative margin audits: investigate any non-zero count
- Margin-floor violators: investigate spikes campaign-over-campaign
- Unmapped inventory: investigate any non-zero count immediately
- Fulfillment fallback: investigate any non-zero count (schema drift risk)

## Logging template

| Run Date | Version | Duration (s) | SKUs | Q1A Cost | Q1A Margin | Q1B Floor | B2B Holds | Unmapped Inventory | Fulfillment Fallbacks | Operator Notes |
|---|---|---:|---:|---:|---:|---:|---:|---:|---:|---|
|  |  |  |  |  |  |  |  |  |  |  |

## Escalation guidance

Escalate to data operations and pricing owner when:

- any mandatory source file is missing
- pre-flight fails on required headers
- guardrail pressure blocks campaign-critical SKU segments
- unmapped inventory or fallback counts persist for 2+ runs
