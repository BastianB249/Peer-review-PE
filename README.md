# Peer-review-PE

Submission-grade peer + WACC workbook pipeline for TKH.

## Repository structure

<<<<<<< HEAD
- `inputs/peer_universe.csv` – peer list, selection flags, optional `gvkey`, segment fit, rationale.
- `inputs/wrds_mapping.csv` – WRDS lookup mapping per ticker (`wrds_db`, `identifier_type`, `identifier_value`).
- `inputs/data_overrides.csv` – optional reproducible manual overrides.
- `scripts/build_peer_model.py` – single entrypoint.
- `outputs/` – local generated artifacts (kept out of git).
=======
- `inputs/peer_universe.csv` – peer list, selection flags, segment fit, and rationale.
- `inputs/data_overrides.csv` – optional manual overrides (kept empty by default).
- `scripts/build_peer_model.py` – single entrypoint that fetches data, builds workbook, and runs QC.
- `outputs/` – local generated artifacts (kept out of git; run build command to regenerate).
>>>>>>> origin/main

## Build command

```bash
python -m scripts.build_peer_model
```

<<<<<<< HEAD
Pipeline order:
1. Parse peers + WRDS mapping.
2. Attempt WRDS pull (default priority when `WRDS_USERNAME` is set and mapping exists).
3. Fill missing fields from Yahoo fallback only when WRDS does not provide those fields.
4. Apply overrides.
5. Build workbook with `Peer_Table`, `WACC_Model`, `Sources_and_AsOf`, `QC_Report`, `Peer_Rationale`, `Clean_Overview`.

## Smoke tests

### Fallback mode (no WRDS)
```bash
unset WRDS_USERNAME
python -m scripts.build_peer_model
```

### WRDS mode (partial mapping still allowed)
```bash
export WRDS_USERNAME=Bastianuser
python -m scripts.build_peer_model
```

If WRDS mappings are incomplete (`identifier_value` blank), those peers automatically fall back and are clearly tagged in `Sources_and_AsOf` and logs.

## Notes

- WRDS credentials must stay outside the repo (`.pgpass` or WRDS auth flow).
- EV mode toggle is in script config (`USE_PROVIDER_EV_AS_TRUTH`).
- Mixed sources are controlled by `ALLOW_MIXED_SOURCES`.
- KPMG ERP/SFP values should be manually confirmed and entered in script assumptions.

## GitHub PR note

Binary Excel outputs are intentionally not committed. Generate `outputs/TKH_Peer_Analysis_submission_ready.xlsx` locally before submission/upload.
=======
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
>>>>>>> origin/main
