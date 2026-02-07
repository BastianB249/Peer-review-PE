from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill


PEERS = [
    ("Aalberts", "AALB.AS", 1, "Closest diversified industrial technology peer"),
    ("ASM International", "ASMI.AS", 1, "Automation/advanced tech valuation anchor"),
    ("Basler", "BSL.DE", 1, "Direct machine vision hardware/software comparable"),
    ("Cognex", "COGX", 1, "Global machine vision leader benchmark"),
    ("Jenoptik", "JEN.DE", 1, "Photonics and optical systems overlap"),
    ("Huber+Suhner", "HUBN.SW", 1, "Connectivity and industrial cabling exposure"),
    ("NKT", "NKT.CO", 1, "Power/subsea cable and electrification exposure"),
    ("Mersen", "MRN.PA", 1, "Electrical components and power management peer"),
    ("Arcadis", "ARCAD.AS", 0, "Services-heavy model, less product/asset intensity"),
    ("Fugro", "FUR.AS", 0, "Geo-data/offshore services, weaker industrial comparability"),
    ("SBM Offshore", "SBMO.AS", 0, "Project/offshore leasing model not close to TKH"),
    ("Vopak", "VPK.AS", 0, "Tank storage infrastructure, limited operating overlap"),
]

BASE_COLUMNS = [
    "Company",
    "Ticker",
    "Selected (1/0)",
    "Selection rationale",
    "Share Price",
    "Market Cap",
    "Enterprise Value",
    "Revenue LTM",
    "EBITDA LTM",
    "EBIT LTM",
    "Net Debt",
]

MULTIPLE_GROUPS = [
    ("EV/Sales", ["2023", "2024"]),
    ("EV/EBITDA", ["2023", "2024"]),
    ("EV/EBIT", ["2023", "2024"]),
]

TAIL_COLUMNS = [
    "Net Debt/EBITDA",
    "EBITDA Margin",
]


def main() -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Peer_Table"

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    row_1 = 1
    row_2 = 2

    current_col = 1
    for label in BASE_COLUMNS:
        sheet.cell(row=row_1, column=current_col, value=label)
        sheet.merge_cells(start_row=row_1, start_column=current_col, end_row=row_2, end_column=current_col)
        current_col += 1

    for group_label, years in MULTIPLE_GROUPS:
        start_col = current_col
        sheet.cell(row=row_1, column=start_col, value=group_label)
        for year in years:
            sheet.cell(row=row_2, column=current_col, value=year)
            current_col += 1
        sheet.merge_cells(start_row=row_1, start_column=start_col, end_row=row_1, end_column=current_col - 1)

    for label in TAIL_COLUMNS:
        sheet.cell(row=row_1, column=current_col, value=label)
        sheet.merge_cells(start_row=row_1, start_column=current_col, end_row=row_2, end_column=current_col)
        current_col += 1

    for row in range(row_1, row_2 + 1):
        for col in range(1, current_col):
            cell = sheet.cell(row=row, column=col)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment

    sheet.freeze_panes = "A3"

    data_start_row = 3
    for offset, (name, ticker, selected, rationale) in enumerate(PEERS):
        row = data_start_row + offset
        sheet.cell(row=row, column=1, value=name)
        sheet.cell(row=row, column=2, value=ticker)
        sheet.cell(row=row, column=3, value=selected)
        sheet.cell(row=row, column=4, value=rationale)

        ev_sales_2024 = f"=IF(H{row}=0,0,G{row}/H{row})"
        ev_ebitda_2024 = f"=IF(I{row}=0,0,G{row}/I{row})"
        ev_ebit_2024 = f"=IF(J{row}=0,0,G{row}/J{row})"
        net_debt_ebitda = f"=IF(I{row}=0,0,K{row}/I{row})"
        ebitda_margin = f"=IF(H{row}=0,0,I{row}/H{row})"

        sheet.cell(row=row, column=13, value=ev_sales_2024)
        sheet.cell(row=row, column=15, value=ev_ebitda_2024)
        sheet.cell(row=row, column=17, value=ev_ebit_2024)
        sheet.cell(row=row, column=18, value=net_debt_ebitda)
        sheet.cell(row=row, column=19, value=ebitda_margin)

    blank_row = data_start_row + len(PEERS)
    summary_start = blank_row + 1
    summary_formulas = [
        ("Selected peers average EV/Sales 2024", f"=AVERAGEIF(C{data_start_row}:C{blank_row - 1},1,M{data_start_row}:M{blank_row - 1})"),
        ("Selected peers median EV/Sales 2024", f"=MEDIAN(IF(C{data_start_row}:C{blank_row - 1}=1,M{data_start_row}:M{blank_row - 1}))"),
        ("Selected peers average EV/EBITDA 2024", f"=AVERAGEIF(C{data_start_row}:C{blank_row - 1},1,O{data_start_row}:O{blank_row - 1})"),
        ("Selected peers median EV/EBITDA 2024", f"=MEDIAN(IF(C{data_start_row}:C{blank_row - 1}=1,O{data_start_row}:O{blank_row - 1}))"),
        ("Selected peers average EV/EBIT 2024", f"=AVERAGEIF(C{data_start_row}:C{blank_row - 1},1,Q{data_start_row}:Q{blank_row - 1})"),
        ("Selected peers median EV/EBIT 2024", f"=MEDIAN(IF(C{data_start_row}:C{blank_row - 1}=1,Q{data_start_row}:Q{blank_row - 1}))"),
    ]

    for offset, (label, formula) in enumerate(summary_formulas):
        row = summary_start + offset
        sheet.cell(row=row, column=1, value=label)
        sheet.cell(row=row, column=2, value=formula)

    column_widths = {
        1: 18,
        2: 12,
        3: 14,
        4: 46,
        5: 12,
        6: 14,
        7: 16,
        8: 14,
        9: 14,
        10: 12,
        11: 12,
        12: 11,
        13: 11,
        14: 11,
        15: 11,
        16: 11,
        17: 11,
        18: 14,
        19: 14,
    }
    for col, width in column_widths.items():
        sheet.column_dimensions[chr(64 + col)].width = width

    instructions = workbook.create_sheet(title="Instructions")
    instructions_data = [
        ["TKH peer workbook", f"Generated on {date.today().isoformat()}"],
        ["How to populate live KPIs", "Paste or link current values for Share Price, Market Cap, EV, Revenue, EBITDA, EBIT, Net Debt in the main sheet."],
        ["Suggested data sources", "Bloomberg, Capital IQ, FactSet, Refinitiv, company filings, or Yahoo Finance."],
        ["Units", "Use one consistent currency/unit (e.g., EURm) for Market Cap/EV/Revenue/EBITDA/EBIT/Net Debt."],
        ["Selection result", "Rows with Selected=1 are the recommended 8 peers for LBO comps."],
    ]
    for row, values in enumerate(instructions_data, start=1):
        for col, value in enumerate(values, start=1):
            instructions.cell(row=row, column=col, value=value)

    workbook.save("TKH_Peer_Analysis.xlsx")


if __name__ == "__main__":
    main()
