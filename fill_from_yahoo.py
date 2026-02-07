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


def _first_value(data: dict[str, Any], keys: list[str]) -> Any:
    for key in keys:
        value = data.get(key)
        if value not in (None, ""):
            return value
    return None


def _to_eurm(value: Any) -> float | None:
    if value in (None, ""):
        return None
    try:
        return float(value) / 1_000_000
    except (TypeError, ValueError):
        return None


def main() -> None:
    workbook = load_workbook(INPUT_FILE)
    sheet = workbook[SHEET_NAME]

    headers = {
        cell.value: cell.column
        for cell in sheet[HEADER_ROW]
        if cell.value
    }

    required_columns = [
        "Ticker",
        "Selected (1/0)",
        "Share Price",
        "Market Cap",
        "Enterprise Value",
        "Revenue LTM",
        "EBITDA LTM",
        "EBIT LTM",
    ]
    missing = [col for col in required_columns if col not in headers]
    if missing:
        raise ValueError(f"Missing required columns: {', '.join(missing)}")

    for row in range(DATA_START_ROW, sheet.max_row + 1):
        ticker = sheet.cell(row=row, column=headers["Ticker"]).value
        selected = sheet.cell(row=row, column=headers["Selected (1/0)"]).value

        if selected != 1 or not ticker:
            continue

        ticker_data = yf.Ticker(str(ticker))
        info = ticker_data.get_info()
        fast_info = getattr(ticker_data, "fast_info", {}) or {}

        share_price = _first_value(
            {**fast_info, **info},
            ["last_price", "regularMarketPrice"],
        )
        market_cap = _first_value(
            {**fast_info, **info},
            ["market_cap", "marketCap"],
        )

        enterprise_value = info.get("enterpriseValue")
        revenue = info.get("totalRevenue")
        ebitda = info.get("ebitda")
        ebit = _first_value(info, ["ebit", "operatingIncome"])

        sheet.cell(row=row, column=headers["Share Price"]).value = share_price
        sheet.cell(row=row, column=headers["Market Cap"]).value = _to_eurm(market_cap)
        sheet.cell(row=row, column=headers["Enterprise Value"]).value = _to_eurm(enterprise_value)
        sheet.cell(row=row, column=headers["Revenue LTM"]).value = _to_eurm(revenue)
        sheet.cell(row=row, column=headers["EBITDA LTM"]).value = _to_eurm(ebitda)

        ebit_cell = sheet.cell(row=row, column=headers["EBIT LTM"])
        ebit_value = _to_eurm(ebit)
        if ebit_value is None:
            ebit_cell.value = None
            ebit_cell.fill = MISSING_EBIT_FILL
        else:
            ebit_cell.value = ebit_value

    workbook.save(OUTPUT_FILE)


if __name__ == "__main__":
    main()
