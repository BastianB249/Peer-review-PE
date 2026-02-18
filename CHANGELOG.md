# Changelog

## 2026-02-18
- Rebuilt the peer analysis process into a reproducible pipeline with a single entry point: `python -m scripts.build_peer_model`.
- Added structured project folders: `inputs/`, `outputs/`, and `scripts/`.
- Added a submission-ready workbook output: `outputs/TKH_Peer_Analysis_submission_ready.xlsx`.
- Added workbook sheets for transparency and control: `Clean_Overview`, `Sources_and_AsOf`, `QC_Report`, and `Peer_Rationale`.
- Added quality-control logic for EV reconciliation, scaling outliers, missing fields/denominator checks, year-over-year consistency, and loss-making flags.
- Upgraded WACC with explicit methodology assumptions, mean + median beta statistics, and median-default headline WACC.

## 2026-02-18 (follow-up)
- Removed generated binary and cache artifacts from version control (`.xlsx`, build log, `__pycache__`).
- Added `.gitignore` rules to prevent re-adding generated outputs/caches.
- Kept `outputs/.gitkeep` so the output directory exists while artifacts are built locally.
- Updated README with guidance for local workbook generation when PR tooling blocks binary files.
