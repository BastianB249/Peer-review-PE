# Peer-review-PE

## Workbook build

Generate the formatted peer analysis workbook template:

```bash
py build_peer_workbook.py
```

## Yahoo Finance fill

Install dependencies and populate peer rows using Yahoo Finance:

```bash
py -m pip install -r requirements.txt
py fill_from_yahoo.py
```

The fill script writes `TKH_Peer_Analysis_filled.xlsx`.
It also supports the TKH Group row (including the `TKH` → `TWEKA.AS` mapping) and fills the TKH Inputs block.
