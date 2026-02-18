from __future__ import annotations

import csv
import logging
import math
import os
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import pandas as pd
import yfinance as yf
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


# ------------------------------
# Configurable assumptions
# ------------------------------
AS_OF_OVERRIDE = None  # e.g., "2026-02-18T10:00:00Z"
FISCAL_YEARS = [2023, 2024]
OUTPUT_FILE = Path("outputs/TKH_Peer_Analysis_submission_ready.xlsx")
PEER_INPUT_FILE = Path("inputs/peer_universe.csv")
OVERRIDE_FILE = Path("inputs/data_overrides.csv")

USE_PROVIDER_EV_AS_TRUTH = True  # Mode 1 true, Mode 2 false (compute EV internally)
INCLUDE_MINORITY_INTEREST = False
INCLUDE_LEASES = False

# WACC assumptions (set from KPMG study if available; source visibility handled in workbook)
RISK_FREE_RATE = 0.030
EQUITY_RISK_PREMIUM = 0.050
SMALL_FIRM_PREMIUM = 0.0125
MARGINAL_TAX_RATE = 0.25
COST_OF_DEBT_PRE_TAX = 0.055
TARGET_D_OVER_E = 0.25
PREFERRED_EQUITY_WEIGHT = 0.0

BETA_HORIZON = "5Y"
BETA_FREQUENCY = "Monthly"
BETA_INDEX = "MSCI World"
COST_OF_DEBT_METHOD = "Assumption (Rf + spread proxy); replace with issuer-implied yield when available"
ERP_SOURCE_NOTE = "KPMG cost of capital study (manual input required; network-restricted in current environment)"
SFP_SOURCE_NOTE = "KPMG size premium table (manual input required; network-restricted in current environment)"

LOG_FILE = Path("outputs/build_peer_model.log")


@dataclass
class PeerRow:
    company: str
    ticker: str
    selected: int
    core_set: int
    segment_fit: str
    peer_status: str
    selection_rationale: str
    currency: str | None = None
    share_price_ccy: float | None = None
    market_cap_ccy_m: float | None = None
    enterprise_value_ccy_m: float | None = None
    gross_debt_ccy_m: float | None = None
    cash_ccy_m: float | None = None
    net_debt_ccy_m: float | None = None
    equity_beta: float | None = None
    fx_to_eur: float | None = None
    revenue: dict[int, float | None] = field(default_factory=dict)
    ebitda: dict[int, float | None] = field(default_factory=dict)
    ebit: dict[int, float | None] = field(default_factory=dict)
    source_market_cap: str = "MISSING SOURCE"
    source_ev: str = "MISSING SOURCE"
    source_net_debt: str = "MISSING SOURCE"
    source_beta: str = "MISSING SOURCE"
    source_financials: str = "MISSING SOURCE"

    @property
    def market_cap_eur_m(self) -> float | None:
        if self.market_cap_ccy_m is None or self.fx_to_eur is None:
            return None
        return self.market_cap_ccy_m * self.fx_to_eur

    @property
    def net_debt_eur_m(self) -> float | None:
        if self.net_debt_ccy_m is None or self.fx_to_eur is None:
            return None
        return self.net_debt_ccy_m * self.fx_to_eur

    @property
    def enterprise_value_input_ccy_m(self) -> float | None:
        if USE_PROVIDER_EV_AS_TRUTH:
            return self.enterprise_value_ccy_m
        if self.market_cap_ccy_m is None or self.net_debt_ccy_m is None:
            return None
        ev = self.market_cap_ccy_m + self.net_debt_ccy_m
        return ev


def setup_logging() -> None:
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[logging.FileHandler(LOG_FILE, mode="w"), logging.StreamHandler()],
    )


def to_m(v: Any) -> float | None:
    if v in (None, ""):
        return None
    try:
        return float(v) / 1_000_000
    except Exception:
        return None


def parse_peers(path: Path) -> list[PeerRow]:
    peers: list[PeerRow] = []
    with path.open("r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            peers.append(
                PeerRow(
                    company=row["company"],
                    ticker=row["ticker"],
                    selected=int(row["selected"]),
                    core_set=int(row["core_set"]),
                    segment_fit=row["segment_fit"],
                    peer_status=row["peer_status"],
                    selection_rationale=row["selection_rationale"],
                    revenue={y: None for y in FISCAL_YEARS},
                    ebitda={y: None for y in FISCAL_YEARS},
                    ebit={y: None for y in FISCAL_YEARS},
                )
            )
    return peers


def _find_row_label(index: list[Any], labels: list[str]) -> Any | None:
    mapping = {str(lbl).lower(): lbl for lbl in index}
    for label in labels:
        if label.lower() in mapping:
            return mapping[label.lower()]
    return None


def _extract_metric_by_year(financials: pd.DataFrame, labels: list[str]) -> dict[int, float]:
    if financials is None or financials.empty:
        return {}
    row_lbl = _find_row_label(list(financials.index), labels)
    if row_lbl is None:
        return {}
    series = financials.loc[row_lbl]
    out: dict[int, float] = {}
    for col, val in series.items():
        if pd.isna(val):
            continue
        try:
            year = pd.Timestamp(col).year
        except Exception:
            continue
        out[year] = float(val)
    return out


def _extract_latest_balance(balance: pd.DataFrame, labels: list[str]) -> float | None:
    if balance is None or balance.empty:
        return None
    row_lbl = _find_row_label(list(balance.index), labels)
    if row_lbl is None:
        return None
    col = balance.columns[0]
    val = balance.loc[row_lbl].get(col)
    if val is None or pd.isna(val):
        return None
    return float(val)


def _last_close(tkr: yf.Ticker) -> float | None:
    try:
        hist = tkr.history(period="5d")
        if hist.empty:
            return None
        return float(hist["Close"].dropna().iloc[-1])
    except Exception:
        return None


def _fetch_fx_rate(ccy: str | None, cache: dict[str, float]) -> float | None:
    if ccy in (None, ""):
        return None
    if ccy == "EUR":
        return 1.0
    if ccy in cache:
        return cache[ccy]
    direct = _last_close(yf.Ticker(f"{ccy}EUR=X"))
    if direct is not None:
        cache[ccy] = direct
        return direct
    inverse = _last_close(yf.Ticker(f"EUR{ccy}=X"))
    if inverse is not None and inverse != 0:
        cache[ccy] = 1.0 / inverse
        return cache[ccy]
    return None


def fetch_market_data(peers: list[PeerRow]) -> None:
    logging.info("Fetching market and financial data from Yahoo Finance (fallback source)...")
    fx_cache: dict[str, float] = {}
    for p in peers:
        try:
            tkr = yf.Ticker(p.ticker)
            info = tkr.get_info() or {}
            fin = tkr.financials
            bal = tkr.balance_sheet
        except Exception as exc:
            logging.warning("%s: data fetch failed: %s", p.ticker, exc)
            continue

        p.currency = info.get("currency")
        p.share_price_ccy = _last_close(tkr)
        p.market_cap_ccy_m = to_m(info.get("marketCap"))
        p.enterprise_value_ccy_m = to_m(info.get("enterpriseValue"))
        p.equity_beta = info.get("beta")

        total_debt = _extract_latest_balance(bal, ["Total Debt", "Long Term Debt", "Current Debt"])
        cash = _extract_latest_balance(
            bal,
            ["Cash And Cash Equivalents", "Cash And Cash Equivalents Including Short Term Investments", "Cash"],
        )
        p.gross_debt_ccy_m = to_m(total_debt)
        p.cash_ccy_m = to_m(cash)
        net_debt = info.get("netDebt")
        p.net_debt_ccy_m = to_m(net_debt) if net_debt is not None else (p.gross_debt_ccy_m - p.cash_ccy_m if p.gross_debt_ccy_m is not None and p.cash_ccy_m is not None else None)

        rev = _extract_metric_by_year(fin, ["Total Revenue", "TotalRevenue"])
        ebitda = _extract_metric_by_year(fin, ["EBITDA"])
        ebit = _extract_metric_by_year(fin, ["EBIT", "Operating Income", "OperatingIncome"])
        if not ebitda:
            da = _extract_metric_by_year(fin, ["Depreciation And Amortization", "Depreciation & Amortization"])
            for y in set(ebit) & set(da):
                ebitda[y] = ebit[y] + da[y]

        for y in FISCAL_YEARS:
            p.revenue[y] = to_m(rev.get(y))
            p.ebitda[y] = to_m(ebitda.get(y))
            p.ebit[y] = to_m(ebit.get(y))

        p.fx_to_eur = _fetch_fx_rate(p.currency, fx_cache)

        source = f"Yahoo Finance ({p.ticker})"
        p.source_market_cap = source if p.market_cap_ccy_m is not None else "MISSING SOURCE"
        p.source_ev = source if p.enterprise_value_ccy_m is not None else "MISSING SOURCE"
        p.source_net_debt = source if p.net_debt_ccy_m is not None else "MISSING SOURCE"
        p.source_beta = source if p.equity_beta is not None else "MISSING SOURCE"
        has_fin = any(v is not None for v in list(p.revenue.values()) + list(p.ebitda.values()) + list(p.ebit.values()))
        p.source_financials = source if has_fin else "MISSING SOURCE"


def try_wrds_enrichment(peers: list[PeerRow]) -> str:
    """Attempt WRDS connection (for reproducibility + traceability); returns status string."""
    username = os.getenv("WRDS_USERNAME")
    if not username:
        return "WRDS not configured (WRDS_USERNAME missing); fallback provider used"
    try:
        import wrds  # type: ignore

        db = wrds.Connection(wrds_username=username)
        db.close()
        # Integration point reserved for institution-specific WRDS SQL mapping.
        return "WRDS connection successful; no mapped query configured, fallback provider retained"
    except Exception as exc:
        return f"WRDS unavailable ({exc}); fallback provider used"


def apply_overrides(peers: list[PeerRow], path: Path) -> None:
    if not path.exists():
        return
    df = pd.read_csv(path)
    if df.empty:
        return
    by_ticker = {p.ticker: p for p in peers}
    for _, row in df.iterrows():
        p = by_ticker.get(str(row.get("ticker", "")))
        if p is None:
            continue
        field_name = str(row.get("field", ""))
        year = row.get("year")
        value = row.get("value")
        if pd.isna(value):
            continue
        if field_name in {"market_cap_ccy_m", "enterprise_value_ccy_m", "net_debt_ccy_m", "equity_beta", "gross_debt_ccy_m", "cash_ccy_m"}:
            setattr(p, field_name, float(value))
        elif field_name in {"revenue", "ebitda", "ebit"} and not pd.isna(year):
            getattr(p, field_name)[int(year)] = float(value)


def metric_multiple(ev: float | None, denom: float | None) -> float | None:
    if ev is None or denom in (None, 0):
        return None
    return ev / denom


def median(values: list[float]) -> float | None:
    clean = sorted(v for v in values if v is not None and not math.isnan(v))
    if not clean:
        return None
    n = len(clean)
    if n % 2 == 1:
        return clean[n // 2]
    return (clean[n // 2 - 1] + clean[n // 2]) / 2


def mean(values: list[float]) -> float | None:
    clean = [v for v in values if v is not None and not math.isnan(v)]
    if not clean:
        return None
    return sum(clean) / len(clean)


def compute_qc_rows(peers: list[PeerRow]) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    for p in peers:
        ev = p.enterprise_value_input_ccy_m
        ev_recon = None if p.market_cap_ccy_m is None or p.net_debt_ccy_m is None else p.market_cap_ccy_m + p.net_debt_ccy_m
        delta = None if ev is None or ev_recon is None else ev - ev_recon
        pct_delta = None if delta is None or ev in (None, 0) else delta / ev
        ev_ok = pct_delta is not None and abs(pct_delta) <= 0.05

        ev_ebitda_2024 = metric_multiple(ev, p.ebitda.get(2024))
        ev_ebit_2024 = metric_multiple(ev, p.ebit.get(2024))
        scale_flag = (ev_ebitda_2024 is not None and abs(ev_ebitda_2024) > 50) or (ev_ebit_2024 is not None and abs(ev_ebit_2024) > 80)

        missing = []
        for field in [p.market_cap_ccy_m, p.enterprise_value_ccy_m, p.net_debt_ccy_m, p.equity_beta]:
            if field is None:
                missing.append("base")
        for y in FISCAL_YEARS:
            for val in [p.revenue[y], p.ebitda[y], p.ebit[y]]:
                if val is None:
                    missing.append(str(y))
        denom_issue = any(p.ebitda[y] in (None, 0) or p.ebit[y] in (None, 0) for y in FISCAL_YEARS)
        loss_year = any((p.ebit[y] or 0) < 0 for y in FISCAL_YEARS)

        consistency_issue = False
        if p.ebitda[2023] not in (None, 0) and p.ebitda[2024] is not None:
            ratio = abs(p.ebitda[2024] / p.ebitda[2023])
            if ratio > 10 or ratio < 0.1:
                consistency_issue = True

        checks = {
            "ev_reconciliation": "PASS" if ev_ok else "FAIL",
            "unit_scaling": "FAIL" if scale_flag else "PASS",
            "missing_or_denominator": "FAIL" if missing or denom_issue else "PASS",
            "year_consistency": "FAIL" if consistency_issue else "PASS",
            "lossmaking": "FLAG" if loss_year else "PASS",
        }
        out.append(
            {
                "company": p.company,
                "ticker": p.ticker,
                "delta_ev": delta,
                "pct_delta_ev": pct_delta,
                "ev_ebitda_2024": ev_ebitda_2024,
                "ev_ebit_2024": ev_ebit_2024,
                "checks": checks,
                "explanation": (
                    "Likely unit mismatch in EBITDA/EBIT" if scale_flag else "No immediate scaling anomaly"
                ),
            }
        )
    return out


def style_header(ws, row: int, start_col: int, end_col: int) -> None:
    fill = PatternFill("solid", fgColor="1F4E78")
    font = Font(color="FFFFFF", bold=True)
    align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col in range(start_col, end_col + 1):
        c = ws.cell(row=row, column=col)
        c.fill = fill
        c.font = font
        c.alignment = align


def build_workbook(peers: list[PeerRow], wrds_status: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Peer_Table"

    years = FISCAL_YEARS
    columns = [
        "Company",
        "Ticker",
        "Selected (1/0)",
        "Core Set (1/0)",
        "Peer Status",
        "Segment Fit",
        "Selection rationale",
        "Currency",
        "Share Price (CCY)",
        "Market Cap (CCY m)",
        "Enterprise Value (CCY m)",
        "Gross Debt (CCY m)",
        "Cash (CCY m)",
        "Net Debt (CCY m)",
        "Equity Beta",
        "FX to EUR",
    ]
    for y in years:
        columns.extend([f"Revenue {y} (CCY m)", f"EBITDA {y} (CCY m)", f"EBIT {y} (CCY m)", f"EV/Sales {y}", f"EV/EBITDA {y}", f"EV/EBIT {y}"])

    for i, col in enumerate(columns, 1):
        ws.cell(row=1, column=i, value=col)
    style_header(ws, 1, 1, len(columns))

    for ridx, p in enumerate(peers, start=2):
        vals: list[Any] = [
            p.company,
            p.ticker,
            p.selected,
            p.core_set,
            p.peer_status,
            p.segment_fit,
            p.selection_rationale,
            p.currency,
            p.share_price_ccy,
            p.market_cap_ccy_m,
            p.enterprise_value_input_ccy_m,
            p.gross_debt_ccy_m,
            p.cash_ccy_m,
            p.net_debt_ccy_m,
            p.equity_beta,
            p.fx_to_eur,
        ]
        ev = p.enterprise_value_input_ccy_m
        for y in years:
            rev = p.revenue[y]
            ebitda = p.ebitda[y]
            ebit = p.ebit[y]
            vals.extend([rev, ebitda, ebit, metric_multiple(ev, rev), metric_multiple(ev, ebitda), metric_multiple(ev, ebit)])
        for cidx, v in enumerate(vals, 1):
            ws.cell(row=ridx, column=cidx, value=v)

    ws.freeze_panes = "A2"
    for c in range(1, len(columns) + 1):
        ws.column_dimensions[get_column_letter(c)].width = 15
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["G"].width = 48

    # Summary blocks: core vs extended
    start = len(peers) + 4
    ws.cell(row=start, column=1, value="Summary Statistics (Selected peers)").font = Font(bold=True)
    ws.cell(row=start + 1, column=1, value="Metric")
    ws.cell(row=start + 1, column=2, value="Core Set Median")
    ws.cell(row=start + 1, column=3, value="Core Set Mean")
    ws.cell(row=start + 1, column=4, value="Extended Set Median")
    ws.cell(row=start + 1, column=5, value="Extended Set Mean")
    style_header(ws, start + 1, 1, 5)

    metrics = [f"EV/Sales {y}" for y in years] + [f"EV/EBITDA {y}" for y in years] + [f"EV/EBIT {y}" for y in years]
    df = pd.DataFrame([
        {
            "selected": p.selected,
            "core": p.core_set,
            **{f"EV/Sales {y}": metric_multiple(p.enterprise_value_input_ccy_m, p.revenue[y]) for y in years},
            **{f"EV/EBITDA {y}": metric_multiple(p.enterprise_value_input_ccy_m, p.ebitda[y]) for y in years},
            **{f"EV/EBIT {y}": metric_multiple(p.enterprise_value_input_ccy_m, p.ebit[y]) for y in years},
        }
        for p in peers
    ])
    core = df[(df["selected"] == 1) & (df["core"] == 1)]
    ext = df[df["selected"] == 1]
    row = start + 2
    for m in metrics:
        ws.cell(row=row, column=1, value=m)
        ws.cell(row=row, column=2, value=median(core[m].dropna().tolist()))
        ws.cell(row=row, column=3, value=mean(core[m].dropna().tolist()))
        ws.cell(row=row, column=4, value=median(ext[m].dropna().tolist()))
        ws.cell(row=row, column=5, value=mean(ext[m].dropna().tolist()))
        row += 1

    # WACC model detailed
    wacc = wb.create_sheet("WACC_Model")
    wacc_headers = ["Company", "Selected (1/0)", "Core Set (1/0)", "Levered Beta", "Net Debt", "Market Cap", "D/E", "Tax Rate", "Unlevered Beta"]
    for i, h in enumerate(wacc_headers, 1):
        wacc.cell(row=1, column=i, value=h)
    style_header(wacc, 1, 1, len(wacc_headers))
    for idx, p in enumerate(peers, 2):
        de = None if p.market_cap_ccy_m in (None, 0) or p.net_debt_ccy_m is None else p.net_debt_ccy_m / p.market_cap_ccy_m
        unlev = None if p.equity_beta is None or de is None else p.equity_beta / (1 + (1 - MARGINAL_TAX_RATE) * de)
        vals = [p.company, p.selected, p.core_set, p.equity_beta, p.net_debt_ccy_m, p.market_cap_ccy_m, de, MARGINAL_TAX_RATE, unlev]
        for c, v in enumerate(vals, 1):
            wacc.cell(row=idx, column=c, value=v)

    sel_unlev = [wacc.cell(row=i, column=9).value for i in range(2, 2 + len(peers)) if wacc.cell(row=i, column=2).value == 1 and wacc.cell(row=i, column=9).value is not None]
    sel_lev = [wacc.cell(row=i, column=4).value for i in range(2, 2 + len(peers)) if wacc.cell(row=i, column=2).value == 1 and wacc.cell(row=i, column=4).value is not None]

    sr = len(peers) + 4
    entries = [
        ("Beta methodology - horizon", BETA_HORIZON),
        ("Beta methodology - frequency", BETA_FREQUENCY),
        ("Beta methodology - index", BETA_INDEX),
        ("Debt definition", "Net debt (gross debt - cash)") ,
        ("EV definition mode", "Provider EV as truth" if USE_PROVIDER_EV_AS_TRUTH else "Computed EV = Market Cap + Net Debt (+ toggles)"),
        ("Mean levered beta (selected)", mean(sel_lev)),
        ("Median levered beta (selected)", median(sel_lev)),
        ("Mean unlevered beta (selected)", mean(sel_unlev)),
        ("Median unlevered beta (selected)", median(sel_unlev)),
    ]
    for i, (k, v) in enumerate(entries, sr):
        wacc.cell(row=i, column=1, value=k)
        wacc.cell(row=i, column=2, value=v)

    med_unlev_row = sr + 8
    calc_start = med_unlev_row + 2
    relevered_beta = (median(sel_unlev) or 0) * (1 + (1 - MARGINAL_TAX_RATE) * TARGET_D_OVER_E)
    cost_equity = RISK_FREE_RATE + relevered_beta * EQUITY_RISK_PREMIUM + SMALL_FIRM_PREMIUM
    cost_debt_after_tax = COST_OF_DEBT_PRE_TAX * (1 - MARGINAL_TAX_RATE)
    debt_weight = TARGET_D_OVER_E / (1 + TARGET_D_OVER_E)
    equity_weight = 1 - debt_weight - PREFERRED_EQUITY_WEIGHT
    wacc_value = equity_weight * cost_equity + debt_weight * cost_debt_after_tax

    calc = [
        ("Risk-free rate", RISK_FREE_RATE),
        ("Market risk premium", EQUITY_RISK_PREMIUM),
        ("Small firm premium", SMALL_FIRM_PREMIUM),
        ("Marginal tax rate", MARGINAL_TAX_RATE),
        ("Cost of debt (pre-tax)", COST_OF_DEBT_PRE_TAX),
        ("Target D/E", TARGET_D_OVER_E),
        ("Relevered beta (median unlevered)", relevered_beta),
        ("Cost of equity (CAPM + SFP)", cost_equity),
        ("Cost of debt (after-tax)", cost_debt_after_tax),
        ("Debt weight", debt_weight),
        ("Equity weight", equity_weight),
        ("WACC", wacc_value),
        ("Cost of debt methodology", COST_OF_DEBT_METHOD),
        ("ERP source note", ERP_SOURCE_NOTE),
        ("Small firm premium source note", SFP_SOURCE_NOTE),
    ]
    wacc.cell(row=calc_start - 1, column=1, value="WACC Calculation").font = Font(bold=True)
    for i, (k, v) in enumerate(calc, calc_start):
        wacc.cell(row=i, column=1, value=k)
        wacc.cell(row=i, column=2, value=v)
    for col in range(1, 10):
        wacc.column_dimensions[get_column_letter(col)].width = 22

    # Sources sheet
    src = wb.create_sheet("Sources_and_AsOf")
    asof = AS_OF_OVERRIDE or datetime.now(timezone.utc).isoformat()
    meta = [
        ("As-of timestamp (UTC)", asof),
        ("Primary provider", "Yahoo Finance (automated fallback)") ,
        ("WRDS status", wrds_status),
        ("FX assumption", "Spot FX (latest close) to EUR"),
        ("EV mode", "Provider EV" if USE_PROVIDER_EV_AS_TRUTH else "Computed EV"),
        ("Include minority interest", INCLUDE_MINORITY_INTEREST),
        ("Include leases", INCLUDE_LEASES),
    ]
    for i, (k, v) in enumerate(meta, 1):
        src.cell(row=i, column=1, value=k)
        src.cell(row=i, column=2, value=v)
    header_row = len(meta) + 2
    hdr = [
        "Company","Ticker","Source: Market Cap","Source: EV","Source: Net Debt / Gross Debt & Cash","Source: Beta","Source: Financials (2023/2024)",
    ]
    for i, h in enumerate(hdr, 1):
        src.cell(row=header_row, column=i, value=h)
    style_header(src, header_row, 1, len(hdr))
    for i, p in enumerate(peers, header_row + 1):
        vals = [p.company, p.ticker, p.source_market_cap, p.source_ev, p.source_net_debt, p.source_beta, p.source_financials]
        for c, v in enumerate(vals, 1):
            src.cell(row=i, column=c, value=v)

    # QC report
    qc = wb.create_sheet("QC_Report")
    qh = [
        "Company","Ticker","EV Reconciliation Status","EV Delta (CCY m)","EV Delta %","Unit/Scaling Status","Missing/Denominator Status","Year Consistency Status","Loss-making Status","2024 EV/EBITDA","2024 EV/EBIT","Explanation",
    ]
    for i, h in enumerate(qh, 1):
        qc.cell(row=1, column=i, value=h)
    style_header(qc, 1, 1, len(qh))
    qc_rows = compute_qc_rows(peers)
    for r, item in enumerate(qc_rows, 2):
        checks = item["checks"]
        vals = [
            item["company"], item["ticker"], checks["ev_reconciliation"], item["delta_ev"], item["pct_delta_ev"],
            checks["unit_scaling"], checks["missing_or_denominator"], checks["year_consistency"], checks["lossmaking"], item["ev_ebitda_2024"], item["ev_ebit_2024"], item["explanation"],
        ]
        for c, v in enumerate(vals, 1):
            qc.cell(row=r, column=c, value=v)

    # Peer rationale
    pr = wb.create_sheet("Peer_Rationale")
    ph = ["Company", "Ticker", "Segment Fit", "Role (Core/Segment-only/Excluded)", "Selected (1/0)", "Rationale"]
    for i, h in enumerate(ph, 1):
        pr.cell(row=1, column=i, value=h)
    style_header(pr, 1, 1, len(ph))
    for r, p in enumerate(peers, 2):
        vals = [p.company, p.ticker, p.segment_fit, p.peer_status, p.selected, p.selection_rationale]
        for c, v in enumerate(vals, 1):
            pr.cell(row=r, column=c, value=v)

    # Clean sheet (selected only; akin to Accell peer + WACC overview)
    clean = wb.create_sheet("Clean_Overview")
    clean.cell(row=1, column=1, value="Weighted Average Cost of Capital").font = Font(bold=True)
    clean.cell(row=1, column=5, value="PEER GROUP (selected only)").font = Font(bold=True)

    wacc_lines = [
        ("Riskfree rate", RISK_FREE_RATE),
        ("Market risk premium", EQUITY_RISK_PREMIUM),
        ("Small firm premium", SMALL_FIRM_PREMIUM),
        ("Cost of debt (pre-tax)", COST_OF_DEBT_PRE_TAX),
        ("Marginal tax rate", MARGINAL_TAX_RATE),
        ("Target D/E", TARGET_D_OVER_E),
        ("Unlevered beta (median)", median(sel_unlev)),
        ("Relevered beta", relevered_beta),
        ("Cost of common equity", cost_equity),
        ("WACC", wacc_value),
    ]
    for i, (k, v) in enumerate(wacc_lines, 3):
        clean.cell(row=i, column=1, value=k)
        clean.cell(row=i, column=2, value=v)

    headers = ["Company", "Levered Beta", "D/E", "Unlevered Beta"]
    for i, h in enumerate(headers, 5):
        clean.cell(row=3, column=i, value=h)
    style_header(clean, 3, 5, 8)
    row = 4
    for p in peers:
        if p.selected != 1:
            continue
        de = None if p.market_cap_ccy_m in (None, 0) or p.net_debt_ccy_m is None else p.net_debt_ccy_m / p.market_cap_ccy_m
        ub = None if p.equity_beta is None or de is None else p.equity_beta / (1 + (1 - MARGINAL_TAX_RATE) * de)
        clean.cell(row=row, column=5, value=p.company)
        clean.cell(row=row, column=6, value=p.equity_beta)
        clean.cell(row=row, column=7, value=de)
        clean.cell(row=row, column=8, value=ub)
        row += 1
    clean.cell(row=row + 1, column=5, value="Mean")
    clean.cell(row=row + 1, column=6, value=mean(sel_lev))
    clean.cell(row=row + 1, column=7, value=mean([x for x in [None if p.market_cap_ccy_m in (None, 0) or p.net_debt_ccy_m is None else p.net_debt_ccy_m / p.market_cap_ccy_m for p in peers if p.selected==1] if x is not None]))
    clean.cell(row=row + 1, column=8, value=mean(sel_unlev))
    clean.cell(row=row + 2, column=5, value="Median")
    clean.cell(row=row + 2, column=6, value=median(sel_lev))
    clean.cell(row=row + 2, column=7, value=median([x for x in [None if p.market_cap_ccy_m in (None, 0) or p.net_debt_ccy_m is None else p.net_debt_ccy_m / p.market_cap_ccy_m for p in peers if p.selected==1] if x is not None]))
    clean.cell(row=row + 2, column=8, value=median(sel_unlev))
    clean.cell(row=row + 4, column=5, value="Note: Headline uses medians from selected peers.")

    for ws_ in wb.worksheets:
        ws_.freeze_panes = ws_.freeze_panes or "A2"

    wb.save(OUTPUT_FILE)
    logging.info("Saved workbook to %s", OUTPUT_FILE)


def main() -> None:
    setup_logging()
    logging.info("Starting peer model build")
    peers = parse_peers(PEER_INPUT_FILE)
    wrds_status = try_wrds_enrichment(peers)
    logging.info(wrds_status)
    fetch_market_data(peers)
    apply_overrides(peers, OVERRIDE_FILE)
    build_workbook(peers, wrds_status)
    logging.info("Build complete")


if __name__ == "__main__":
    main()
