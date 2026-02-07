# Peer-review-PE

## Workbook build

Generate the formatted peer analysis workbook template:

```bash
py build_peer_workbook.py
```

## Yahoo Finance fill

Install dependencies and populate selected peer rows using Yahoo Finance:

```bash
py -m pip install -r requirements.txt
py fill_from_yahoo.py
```
