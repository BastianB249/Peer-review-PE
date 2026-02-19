# Changelog

## 2026-02-18 (WRDS integration fix)
- Replaced placeholder WRDS connectivity-only logic with real WRDS fundamentals pull in `scripts/build_peer_model.py`.
- Added `fetch_from_wrds(...)` that queries Compustat North America (`comp.funda`) or Compustat Global (`compg.g_funda`) by mapped `gvkey` for FY 2023/2024.
- Implemented deterministic row selection when multiple WRDS rows exist (latest `datadate` per `fyear` after statement filters).
- Added robust WRDS mapping file `inputs/wrds_mapping.csv` with required fields (`ticker, region, wrds_db, identifier_type, identifier_value, notes`).
- Updated `inputs/peer_universe.csv` to include optional `gvkey` column for reproducibility.
- Made WRDS default source when configured; Yahoo now fills market fields and only missing WRDS fields (if `ALLOW_MIXED_SOURCES=True`).
- Extended `Sources_and_AsOf` with WRDS pull status section (connected, mapping coverage, per-peer status) and per-field source tracking.
- Preserved QC behavior (no silent auto-fixes); missing/scaling/reconciliation issues remain visible in `QC_Report`.
- Updated README with fallback and WRDS smoke-test commands.
