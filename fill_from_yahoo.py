from __future__ import annotations

from typing import Any, Dict, Iterable

import pandas as pd
import yfinance as yf
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

INPUT_FILE = "TKH_Peer_Analysis.xlsx"
OUTPUT_FILE = "TKH_Peer_Analysis_filled.xlsx"
SHEET_NAME = "Peer_Table"

HEADER_ROW_1 = 1
HEADER_ROW_2 = 2
DATA_START_ROW = 3

MISSING_OPERATING_FILL = PatternFill("solid", fgColor="FFF2CC")

# Map “your” tickers -> Yahoo tickers when needed
TICKER_MAP = {
    "COGX": "CGNX",        # Cognex
    "ASMI.AS": "ASM.AS",   # ASM International
    # Add TKH mapping only if your sheet uses "TKH" as ticker.
    # If your sheet uses the correct Yahoo symbol already (e.g. TWEKA.AS), you don't need this.
    "TKH": "TWEKA.AS",     # TKH Group
}

REVENUE_LABELS = ["Total Revenue", "TotalRevenue"]
EBIT_LABELS = ["EBIT", "Operating Income", "OperatingIncome"]
EBITDA_LABELS = ["EBITDA"]
DA_LABELS = [
    "Depreciation & Amortization",
    "Depreciation And Amortization",
    "Depreciation/Amortization",
]


def _map_ticker(ticker: str) -> str:
    return TICKER_MAP.get(ticker.strip(), ticker.strip())


def _to_ccy_m(value: Any) -> float | None:
    """Convert absolute currency units -> currency millions."""
    if value in (None, ""):
        return None
    try:
        return float(value) / 1_000_000
    except (TypeError, ValueError):
        return None


def _last_close_price(tkr: yf.Ticker) -> float | None:
    """Get last close price from recent trading days (more robust than info)."""
    try:
        hist = tkr.history(period="5d")
        if hist is None or hist.empty:
            return None
        close = hist["Close"].dropna()
        if close.empty:
            return None
        return float(close.iloc[-1])
    except Exception:
        return None


def _find_row_label(index: Iterable[Any], labels: Iterable[str]) -> Any | None:
    """Find best-matching row label in a yfinance financials dataframe index."""
    label_map = {str(label).lower(): label for label in index}
    for candidate in labels:
        key = candidate.lower()
        if key in label_map:
            return label_map[key]
    return None


def _extract_metric_by_year(financials: pd.DataFrame, labels: Iterable[str]) -> Dict[int, float]:
    """Return {year: value} for a given metric row from tkr.financials."""
    if financials is None or financials.empty:
        return {}
    row_label = _find_row_label(financials.index, labels)
    if row_label is None:
        return {}
    series = financials.loc[row_label]

    data: Dict[int, float] = {}
    for col, value in series.items():
        if pd.isna(value):
            continue
        try:
            year = pd.Timestamp(col).year
        except Exception:
            try:
                year = int(str(col)[:4])
            except Exception:
                continue
        try:
            data[int(year)] = float(value)
        except Exception:
            continue
    return data


def _extract_years(group_map: Dict[tuple[str, str], int], label: str) -> list[int]:
    """Read available years from the 2-row header group map (e.g. Revenue (CCY m) / 2023, 2024)."""
    years: list[int] = []
    for (group_label, year_label) in group_map:
        if group_label == label:
            try:
                years.append(int(year_label))
            except (TypeError, ValueError):
                continue
    return sorted(years)


def _build_header_maps(ws) -> tuple[Dict[str, int], Dict[tuple[str, str], int]]:
    """
    Build:
      base_cols: { "Ticker": col, ... } for columns whose row2 header is blank
      group_cols: { ("Revenue (CCY m)", "2023"): col, ... } for grouped year columns

    Handles merged cells in row1 where only the leftmost cell contains the group label.
    """
    base_cols: Dict[str, int] = {}
    group_cols: Dict[tuple[str, str], int] = {}

    last_header_1: str | None = None

    for col in range(1, ws.max_column + 1):
        header_1 = ws.cell(row=HEADER_ROW_1, column=col).value
        header_2 = ws.cell(row=HEADER_ROW_2, column=col).value

        if header_1 is not None and str(header_1).strip() != "":
            last_header_1 = str(header_1).strip()

        if last_header_1 is None:
            continue

        if header_2 is None or str(header_2).strip() == "":
            base_cols[last_header_1] = col
        else:
            group_cols[(last_header_1, str(header_2).strip())] = col

    return base_cols, group_cols


def _write_operating_value(ws, row: int, col: int, value: float | None) -> None:
    cell = ws.cell(row=row, column=col)
    if value is None:
        cell.value = None
        cell.fill = MISSING_OPERATING_FILL
    else:
        cell.value = _to_ccy_m(value)


def main() -> None:
    wb = load_workbook(INPUT_FILE)
    ws = wb[SHEET_NAME]

    base_cols, group_cols = _build_header_maps(ws)

    required_base = [
        "Ticker",
        "Currency",
        "Share Price (CCY)",
        "Market Cap (CCY m)",
        "Enterprise Value (CCY m)",
        "Net Debt (CCY m)",
    ]
    missing = [c for c in required_base if c not in base_cols]
    if missing:
        raise ValueError(f"Missing required base columns (row {HEADER_ROW_1}): {missing}")

    years = _extract_years(group_cols, "Revenue (CCY m)")
    if len(years) < 2:
        raise ValueError("Expected two fiscal years in the header for Revenue (CCY m).")

    for y in years:
        for group in ("Revenue (CCY m)", "EBITDA (CCY m)", "EBIT (CCY m)"):
            if (group, str(y)) not in group_cols:
                raise ValueError(f"Missing grouped column: {group} / {y}")

    selected_col_exists = "Selected (1/0)" in base_cols

    for row in range(DATA_START_ROW, ws.max_row + 1):
        ticker_val = ws.cell(row=row, column=base_cols["Ticker"]).value
        if not ticker_val:
            continue

        if selected_col_exists:
            selected_val = ws.cell(row=row, column=base_cols["Selected (1/0)"]).value
            if str(selected_val).strip() != "1":
                continue

        raw = str(ticker_val).strip()
        ysym = _map_ticker(raw)

        try:
            tkr = yf.Ticker(ysym)
        except Exception as exc:
            print(f"{raw} -> {ysym}: ticker init failed: {exc}")
            continue

        share_price = _last_close_price(tkr)

        try:
            info = tkr.get_info() or {}
        except Exception as exc:
            print(f"{raw} -> {ysym}: get_info failed: {exc}")
            info = {}

        currency = info.get("currency")
        market_cap = info.get("marketCap")
        enterprise_value = info.get("enterpriseValue")
        net_debt = info.get("netDebt")

        try:
            financials = tkr.financials
        except Exception as exc:
            print(f"{raw} -> {ysym}: financials failed: {exc}")
            financials = pd.DataFrame()

        revenue_by_year = _extract_metric_by_year(financials, REVENUE_LABELS)
        ebit_by_year = _extract_metric_by_year(financials, EBIT_LABELS)
        ebitda_by_year = _extract_metric_by_year(financials, EBITDA_LABELS)

        # If EBITDA missing, try EBIT + D&A
        if not ebitda_by_year:
            da_by_year = _extract_metric_by_year(financials, DA_LABELS)
            for year in set(ebit_by_year) & set(da_by_year):
                ebitda_by_year[year] = ebit_by_year[year] + da_by_year[year]

        # Write base cells
        ws.cell(row=row, column=base_cols["Currency"], value=currency)
        ws.cell(row=row, column=base_cols["Share Price (CCY)"], value=share_price)
        ws.cell(row=row, column=base_cols["Market Cap (CCY m)"], value=_to_ccy_m(market_cap))
        ws.cell(row=row, column=base_cols["Enterprise Value (CCY m)"], value=_to_ccy_m(enterprise_value))
        ws.cell(row=row, column=base_cols["Net Debt (CCY m)"], value=_to_ccy_m(net_debt))

        # Write year-group operating metrics
        for year in years:
            _write_operating_value(ws, row, group_cols[("Revenue (CCY m)", str(year))], revenue_by_year.get(year))
            _write_operating_value(ws, row, group_cols[("EBITDA (CCY m)", str(year))], ebitda_by_year.get(year))
            _write_operating_value(ws, row, group_cols[("EBIT (CCY m)", str(year))], ebit_by_year.get(year))

        print(f"Filled {raw} (Yahoo: {ysym})")

    wb.save(OUTPUT_FILE)
    print(f"Saved {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
