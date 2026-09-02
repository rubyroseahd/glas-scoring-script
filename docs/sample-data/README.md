# Sample Data Pack (Sanitized)

This folder provides minimal sanitized CSV fixtures for validating ingestion and header mapping.

## Included files

- `shopify_export_gt.csv`
- `Total sales by product.csv`
- `Cost_Data.csv`
- `EEI WEB Whse Stock Report.csv`
- `EEI USA Whse Stock Report.csv`

## How to use

1. Copy these files into the Google Drive folder configured via Script Property `VDM_FOLDER_ID` (preferred), or `VDM_CONFIG.FOLDER_ID` in `/home/runner/work/glas-scoring-script/glas-scoring-script/Config.gs` as a legacy fallback.
2. Keep file names exactly as provided.
3. Run **EEI Pricing Engine Launcher → Advanced Diagnostics → Run Pre-Flight Sanity Check**.
4. Run **EEI Pricing Engine Launcher → 1. Full System Sync (Standard)**.

## Notes

- Values are synthetic and non-sensitive.
- EEI warehouse samples intentionally preserve rows 1–4 metadata and row 5 headers.
