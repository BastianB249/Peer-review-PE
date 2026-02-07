from __future__ import annotations

from typing import Any

import yfinance as yf
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

INPUT_FILE = "TKH_Peer_Analysis.xlsx"
OUTPUT_FILE = "TKH_Peer_Analysis_filled.xlsx"
SHEET_NAME = "Peer_Table"
HEADER_ROW = 1
DATA_START_ROW = 3

MISSING_EBIT_FILL = PatternFill("solid", fgColor="FFF2CC")

TICKER_MAP = {
    "COGX": "CGNX",     # Cognex on Yahoo
    "ASMI.AS": "ASM.AS" # ASM International on Yahoo
}



def _map_ticker(t: str) -> str:
    return TICKER_MAP.get(t.strip(), t.strip())


def _to_eurm(value: Any) -> float | None:
    if value in (None, ""):
        return None
    try:
        return float(value) / 1_000_000
    except (TypeError, ValueError):
        return None


def _last_close_price(tkr: yf.Ticker) -> float | None:
    # Use price history instead of fast_info (more stable)
    try:
        hist = tkr.history(period="5d")
        if hist is None or hist.empty:
            return None
        return float(hist["Close"].dropna().iloc[-1])
    except Exception:
        return None


def main() -> None:
    wb = load_workbook(INPUT_FILE)
    ws = wb[SHEET_NAME]

    headers = {cell.value: cell.column for cell in ws[HEADER_ROW] if cell.value}

    required = [
        "Ticker",
        "Selected (1/0)",
        "Share Price",
        "Market Cap",
        "Enterprise Value",
        "Revenue LTM",
        "EBITDA LTM",
        "EBIT LTM",
    ]
    missing = [c for c in required if c not in headers]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    for row in range(DATA_START_ROW, ws.max_row + 1):
        ticker = ws.cell(row=row, column=headers["Ticker"]).value
        selected = ws.cell(row=row, column=headers["Selected (1/0)"]).value

        if not ticker or str(selected).strip() != "1":
            continue

        raw = str(ticker).strip()
        ysym = _map_ticker(raw)

        tkr = yf.Ticker(ysym)

        # Price from history (stable)
        share_price = _last_close_price(tkr)

        # Fundamentals from info (may be incomplete; don't crash)
        try:
            info = tkr.get_info() or {}
        except Exception as e:
            print(f"{raw} -> {ysym}: get_info failed: {e}")
            info = {}

        market_cap = info.get("marketCap")
        enterprise_value = info.get("enterpriseValue")
        revenue = info.get("totalRevenue")
        ebitda = info.get("ebitda")
        ebit = info.get("ebit") or info.get("operatingIncome")

        ws.cell(row=row, column=headers["Share Price"]).value = share_price
        ws.cell(row=row, column=headers["Market Cap"]).value = _to_eurm(market_cap)
        ws.cell(row=row, column=headers["Enterprise Value"]).value = _to_eurm(enterprise_value)
        ws.cell(row=row, column=headers["Revenue LTM"]).value = _to_eurm(revenue)
        ws.cell(row=row, column=headers["EBITDA LTM"]).value = _to_eurm(ebitda)

        ebit_cell = ws.cell(row=row, column=headers["EBIT LTM"])
        ebit_value = _to_eurm(ebit)
        if ebit_value is None:
            ebit_cell.value = None
            ebit_cell.fill = MISSING_EBIT_FILL
        else:
            ebit_cell.value = ebit_value

        print(f"Filled {raw} (Yahoo: {ysym}) price={share_price}")

    wb.save(OUTPUT_FILE)
    print(f"Saved {OUTPUT_FILE}")


if __name__ == "__main__":
    main()

