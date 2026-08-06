# Expected Output Contract

This contract defines key tabs and minimum columns expected after a successful full sync.

## [02] Dashboard Matrix (required)

Minimum critical columns:

- `SKU Anchor Key`
- `Gatekeeper Status`
- `Gatekeeper Code`
- `Fulfillment Tag`
- `Resolved Cost Base`
- `Live Storefront Price`
- `Current Gross Margin %`
- `Raw 90D Retail Velocity`
- `Retail Velocity Score Component`
- `Margin Score Component`
- `Retail Stock Score Component`
- `Total Composite Score`
- `Target Strategic Tier`
- `Tier Code`
- `VDM Markdown Depth %`
- `New Proposed Storefront Price`
- `Final Simulated Stacked Margin %`
- `Profit Guardrail Status Alert`
- `Guardrail Code`
- `Pricing Migration Status`
- `Action Queue`
- `Queue Code`

## [00] Action Items & Sign-off (required)

Minimum critical columns:

- `Queue`
- `SKU`
- `Handle`
- `Vendor`
- `Fulfillment`
- `Total Score`
- `Current Margin %`
- `Stacked Margin %`
- `Guardrail`
- `Target Tier`
- `Proposed Price`
- `Action Required`

## [07] Storefront Update & Sync Audit (required)

Minimum critical columns:

- `SKU Key`
- `Handle`
- `Action`
- `Strategy Tier`
- `VDM Markdown %`
- `Old Price`
- `New Price`
- `Old MSRP`
- `New MSRP`
- `Base Price`
- `Guardrail`

## [09] Master Pricing & Margin Ledger (required)

Minimum critical columns:

- `SKU Key`
- `Handle`
- `Fulfillment`
- `Gatekeeper`
- `Migration Status`
- `Tier`
- `Target Mkdn %`
- `Old Price`
- `New Price`
- `Price Shift ($)`
- `Procurement Cost`
- `Checkout Price`
- `Stacked Margin %`
- `Guardrail`
- `Margin Shift %`

## Validation checklist

After full sync:

- [ ] `[02] Dashboard Matrix` exists and has rows for active Shopify SKUs
- [ ] `[00] Action Items & Sign-off` is regenerated
- [ ] `[07] Storefront Update & Sync Audit` is regenerated
- [ ] `[09] Master Pricing & Margin Ledger` is regenerated
- [ ] `_backup_matrix_data` is refreshed
