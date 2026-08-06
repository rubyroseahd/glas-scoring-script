# Changelog

All notable changes to scoring thresholds, guardrails, queue routing, and operational contracts should be documented here.

## [3.0.0-PROD] - 2026-08-06

### Added

- Operator quick-start workflow in README.
- Sanitized sample-data pack for ingestion/header validation.
- Monthly UAT checklist template with pass/fail evidence fields.
- Failure playbooks for common ingestion and guardrail issues.
- Expected output contract for key tabs and required columns.
- Monitoring metrics guide for run-health tracking.
- Lightweight in-script regression harness entry point (`runRegressionHarness`).

### Notes for future entries

For each release, record:

- Threshold changes (velocity, margin, stock, or tier boundaries)
- Guardrail logic updates (`ERR_*`, `WARN_*` behavior)
- Queue precedence or routing changes
- New/removed required source fields or tabs
