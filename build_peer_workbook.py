from datetime import date
import xml.etree.ElementTree as ET

peers = [
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

header = [
    "Company", "Ticker", "Selected (1/0)", "Selection rationale",
    "Share Price", "Market Cap", "Enterprise Value", "Revenue LTM", "EBITDA LTM", "EBIT LTM", "Net Debt",
    "EV/Sales", "EV/EBITDA", "EV/EBIT", "Net Debt/EBITDA", "EBITDA Margin"
]


def esc(text: str) -> str:
    return (str(text)
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('"', '&quot;'))


def cell(value: str, value_type: str = "String", formula: str | None = None) -> str:
    formula_attr = f' ss:Formula="{esc(formula)}"' if formula else ""
    return f'<Cell{formula_attr}><Data ss:Type="{value_type}">{esc(value)}</Data></Cell>'


rows = ['<Row>' + ''.join(cell(h) for h in header) + '</Row>']

for name, ticker, selected, rationale in peers:
    rows.append('<Row>' + ''.join([
        cell(name),
        cell(ticker),
        cell(selected, "Number"),
        cell(rationale),
        cell(''), cell(''), cell(''), cell(''), cell(''), cell(''), cell(''),
        cell('', formula='=IF(RC[-4]=0,0,RC[-5]/RC[-4])'),
        cell('', formula='=IF(RC[-4]=0,0,RC[-6]/RC[-4])'),
        cell('', formula='=IF(RC[-4]=0,0,RC[-7]/RC[-4])'),
        cell('', formula='=IF(RC[-6]=0,0,RC[-4]/RC[-6])'),
        cell('', formula='=IF(RC[-8]=0,0,RC[-7]/RC[-8])')
    ]) + '</Row>')

rows.append('<Row></Row>')

summary_formulas = [
    ("Selected peers average EV/Sales", "=AVERAGEIF(R2C3:R13C3,1,R2C12:R13C12)"),
    ("Selected peers median EV/Sales", "=MEDIAN(IF(R2C3:R13C3=1,R2C12:R13C12))"),
    ("Selected peers average EV/EBITDA", "=AVERAGEIF(R2C3:R13C3,1,R2C13:R13C13)"),
    ("Selected peers median EV/EBITDA", "=MEDIAN(IF(R2C3:R13C3=1,R2C13:R13C13))"),
    ("Selected peers average EV/EBIT", "=AVERAGEIF(R2C3:R13C3,1,R2C14:R13C14)"),
    ("Selected peers median EV/EBIT", "=MEDIAN(IF(R2C3:R13C3=1,R2C14:R13C14))"),
]

for label, formula in summary_formulas:
    rows.append('<Row>' + cell(label) + cell('', formula=formula) + '</Row>')

instructions = [
    ["TKH peer workbook", f"Generated on {date.today().isoformat()}"],
    ["How to populate live KPIs", "Paste or link current values for Share Price, Market Cap, EV, Revenue, EBITDA, EBIT, Net Debt in the main sheet."],
    ["Suggested data sources", "Bloomberg, Capital IQ, FactSet, Refinitiv, company filings, or Yahoo Finance."],
    ["Units", "Use one consistent currency/unit (e.g., EURm) for Market Cap/EV/Revenue/EBITDA/EBIT/Net Debt."],
    ["Selection result", "Rows with Selected=1 are the recommended 8 peers for LBO comps."],
]
inst_rows = ['<Row>' + ''.join(cell(value) for value in row) + '</Row>' for row in instructions]

xml = f'''<?xml version="1.0"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
 <Worksheet ss:Name="Peer_Table"><Table>{''.join(rows)}</Table></Worksheet>
 <Worksheet ss:Name="Instructions"><Table>{''.join(inst_rows)}</Table></Worksheet>
</Workbook>'''

with open('TKH_Peer_Analysis.xml', 'w', encoding='utf-8') as output:
    output.write(xml)

ET.parse('TKH_Peer_Analysis.xml')
print('Generated valid TKH_Peer_Analysis.xml')
