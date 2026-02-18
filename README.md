# Peer-review-PE

Submission-grade peer + WACC workbook pipeline for TKH.

## Repository structure

- `inputs/peer_universe.csv` – peer list, selection flags, segment fit, and rationale.
- `inputs/data_overrides.csv` – optional manual overrides (kept empty by default).
- `scripts/build_peer_model.py` – single entrypoint that fetches data, builds workbook, and runs QC.
- `outputs/` – local generated artifacts (kept out of git; run build command to regenerate).

## Build command

```bash
python -m scripts.build_peer_model
```

This command:
1. Loads peer inputs.
2. Attempts WRDS connectivity (if `WRDS_USERNAME` is configured).
3. Pulls market + financial data from fallback provider where needed.
4. Writes the workbook with:
   - `Clean_Overview`
   - `Peer_Table`
   - `WACC_Model`
   - `Sources_and_AsOf`
   - `QC_Report`
   - `Peer_Rationale`
5. Saves logs for reproducibility.

## Notes

- EV reconciliation supports two modes in `scripts/build_peer_model.py`:
  - `USE_PROVIDER_EV_AS_TRUTH = True`
  - `USE_PROVIDER_EV_AS_TRUTH = False` (compute EV internally from Market Cap + Net Debt)
- If source fields are unavailable, the model writes explicit `MISSING SOURCE` flags.
- KPMG-based ERP/SFP assumptions are surfaced with source-note cells and are intended for manual confirmation if network access is restricted.


## GitHub PR note

Binary Excel outputs are intentionally not committed.
Generate `outputs/TKH_Peer_Analysis_submission_ready.xlsx` locally before submission/upload.
