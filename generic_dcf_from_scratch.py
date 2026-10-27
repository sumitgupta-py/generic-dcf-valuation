from __future__ import annotations

import argparse
import math
import re
from pathlib import Path
from typing import Dict, Optional, Tuple

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.dimensions import ColumnDimension

NAVY = "17365D"
DARK_BLUE = "1F4E78"
MID_BLUE = "5B9BD5"
LIGHT_BLUE = "D9EAF7"
PALE_BLUE = "EAF2F8"
INPUT_YELLOW = "FFF2CC"
GREEN = "E2F0D9"
GREY = "F2F2F2"
WHITE = "FFFFFF"
BLACK = "222222"
RED = "FCE4D6"
MILLIONS_FMT = '#,##0.0,,;[Red](#,##0.0,,);-'
MILLIONS_FMT_2DP = '#,##0.00,,;[Red](#,##0.00,,);-'
PER_SHARE_FMT = '#,##0.00;[Red](#,##0.00);-'
THIN_GREY = Side(style="thin", color="D9E1F2")
MED_BLUE = Side(style="medium", color=NAVY)

LINE_ITEMS = [
    "Revenue", "Cost of Revenue", "Gross Profit", "Operating Expenses",
    "EBITDA", "Depreciation & Amortization", "EBIT", "Interest Expense",
    "Pretax Income", "Income Tax Expense", "Net Income", "Cash & Cash Equivalents",
    "Total Debt", "Shares Outstanding", "Capital Expenditures",
    "Change in Net Working Capital"
]

ALIASES = {
    "Revenue": ["Total Revenue", "Revenue", "Operating Revenue", "Sales"],
    "Cost of Revenue": ["Cost Of Revenue", "Cost of Revenue", "Cost Of Goods Sold", "Cost of Goods Sold"],
    "Gross Profit": ["Gross Profit"],
    "Operating Expenses": ["Operating Expense", "Operating Expenses", "Total Operating Expenses"],
    "EBITDA": ["EBITDA", "Normalized EBITDA"],
    "Depreciation & Amortization": ["Reconciled Depreciation", "Depreciation And Amortization", "Depreciation & Amortization", "Depreciation"],
    "EBIT": ["Operating Income", "EBIT"],
    "Interest Expense": ["Interest Expense", "Interest Expense Non Operating", "Net Interest Income"],
    "Pretax Income": ["Pretax Income", "Pretax Income", "Income Before Tax"],
    "Income Tax Expense": ["Tax Provision", "Income Tax Expense"],
    "Net Income": ["Net Income", "Net Income Common Stockholders", "Net Income From Continuing Operation Net Minority Interest"],
    "Cash & Cash Equivalents": ["Cash Cash Equivalents And Short Term Investments", "Cash And Cash Equivalents", "Cash"],
    "Total Debt": ["Total Debt", "Long Term Debt And Capital Lease Obligation", "Long Term Debt"],
    "Shares Outstanding": ["Ordinary Shares Number", "Share Issued", "Shares Outstanding"],
    "Capital Expenditures": ["Capital Expenditure", "Capital Expenditures", "Purchase Of PPE"],
    "Change in Net Working Capital": ["Changes In Working Capital", "Change In Working Capital", "Working Capital"]
}

def clean_label(value) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(value).strip().lower()).strip()

def as_number(value, default=0.0) -> float:
    try:
        if pd.isna(value):
            return default
        return float(value)
    except (TypeError, ValueError):
        return default

def find_row(frame: pd.DataFrame, aliases) -> Optional[pd.Series]:
    if frame is None or frame.empty:
        return None
    normalized = {clean_label(index): index for index in frame.index}
    for alias in aliases:
        key = clean_label(alias)
        if key in normalized:
            return frame.loc[normalized[key]]
    for alias in aliases:
        key = clean_label(alias)
        for norm, original in normalized.items():
            if key and (key in norm or norm in key):
                return frame.loc[original]
    return None

def load_yfinance_financials(ticker: str) -> Tuple[pd.DataFrame, str, float, float]:
    try:
        import yfinance as yf
    except ImportError as exc:
        raise RuntimeError("yfinance is not installed. Run: pip install yfinance") from exc
    company = yf.Ticker(ticker.strip())
    income = company.income_stmt
    balance = company.balance_sheet
    cashflow = company.cashflow
    if income is None or income.empty:
        raise RuntimeError(f"No annual income statement returned for {ticker}. Use an uploaded CSV/XLSX instead.")
    years = []
    for col in income.columns:
        try:
            years.append(pd.Timestamp(col).year)
        except Exception:
            continue
    years = sorted(set(years), reverse=True)[:5]
    records = []
    for year in sorted(years):
        col = next((c for c in income.columns if pd.Timestamp(c).year == year), None)
        if col is None:
            continue
        row = {"Fiscal Year": year}
        for item in LINE_ITEMS:
            row[item] = None
        for item, aliases in ALIASES.items():
            source = income
            if item in ["Cash & Cash Equivalents", "Total Debt", "Shares Outstanding"]:
                source = balance
            elif item in ["Capital Expenditures", "Change in Net Working Capital"]:
                source = cashflow
            series = find_row(source, aliases)
            if series is not None and col in series.index:
                value = as_number(series[col], None)
                if item == "Capital Expenditures" and value is not None:
                    value = abs(value)
                if item == "Change in Net Working Capital" and value is not None:
                    value = abs(value)
                row[item] = value
        if row["Revenue"] is not None:
            records.append(row)
    if not records:
        raise RuntimeError(f"Could not map revenue from yfinance for {ticker}. Upload historical financials.")
    frame = pd.DataFrame(records).sort_values("Fiscal Year").reset_index(drop=True)
    info = {}
    try:
        info = company.info or {}
    except Exception:
        pass
    price = as_number(info.get("currentPrice") or info.get("regularMarketPrice"), 0)
    shares = as_number(info.get("sharesOutstanding"), 0)
    if shares <= 0 and "Shares Outstanding" in frame:
        shares = as_number(frame["Shares Outstanding"].dropna().iloc[-1], 0) if frame["Shares Outstanding"].notna().any() else 0
    return frame, ticker.upper(), price, shares

def load_uploaded_financials(path: str) -> Tuple[pd.DataFrame, str, float, float]:
    file_path = Path(path).expanduser().resolve()
    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")
    if file_path.suffix.lower() == ".csv":
        raw = pd.read_csv(file_path)
    elif file_path.suffix.lower() in [".xlsx", ".xlsm"]:
        raw = pd.read_excel(file_path, sheet_name=0)
    else:
        raise ValueError("Upload must be .csv, .xlsx, or .xlsm")
    if raw.empty:
        raise ValueError("Uploaded file has no data.")
    raw.columns = [str(c).strip() for c in raw.columns]
    first_col = raw.columns[0]
    if clean_label(first_col) in ["metric", "line item", "financial metric", "item"]:
        raw = raw.set_index(first_col).T
        raw.index.name = "Fiscal Year"
        raw = raw.reset_index()
    year_col = next((c for c in raw.columns if clean_label(c) in ["fiscal year", "year", "period", "date"]), raw.columns[0])
    records = []
    for _, row in raw.iterrows():
        year_value = row[year_col]
        try:
            year = pd.Timestamp(year_value).year if not str(year_value).isdigit() else int(float(year_value))
        except Exception:
            try:
                year = int(float(year_value))
            except Exception:
                continue
        rec = {"Fiscal Year": year}
        for item in LINE_ITEMS:
            rec[item] = None
        for item, aliases in ALIASES.items():
            match = next((c for c in raw.columns if clean_label(c) in [clean_label(a) for a in aliases]), None)
            if match is not None:
                rec[item] = as_number(row[match], None)
        records.append(rec)
    frame = pd.DataFrame(records)
    if frame.empty or "Revenue" not in frame or frame["Revenue"].notna().sum() == 0:
        raise ValueError("Could not identify Revenue. Use one row per fiscal year and a column named Revenue.")
    frame = frame.sort_values("Fiscal Year").drop_duplicates("Fiscal Year", keep="last").tail(5).reset_index(drop=True)
    for col in LINE_ITEMS:
        if col not in frame:
            frame[col] = None
    return frame[["Fiscal Year"] + LINE_ITEMS], file_path.stem.upper(), 0.0, 0.0

def ask_float(label: str, default: float, low=None, high=None) -> float:
    while True:
        raw = input(f"{label} [{default:.2%}]: ").strip()
        if not raw:
            value = default
        else:
            try:
                value = float(raw.replace("%", "")) / (100 if "%" in raw else 1)
            except ValueError:
                print("Enter a number, for example 12% or 0.12.")
                continue
        if low is not None and value < low:
            print(f"Value must be at least {low}.")
            continue
        if high is not None and value > high:
            print(f"Value must be no more than {high}.")
            continue
        return value

def ask_number(label: str, default: float, minimum=None) -> float:
    while True:
        raw = input(f"{label} [{default:,.2f}]: ").strip()
        if not raw:
            value = default
        else:
            try:
                value = float(raw.replace(",", ""))
            except ValueError:
                print("Enter a valid number.")
                continue
        if minimum is not None and value < minimum:
            print(f"Value must be at least {minimum}.")
            continue
        return value

def ask_assumptions(frame: pd.DataFrame, company: str, current_price: float, shares: float) -> dict:
    last = frame.iloc[-1]
    rev = as_number(last.get("Revenue"), 0)
    ebitda = as_number(last.get("EBITDA"), 0)
    margin = ebitda / rev if rev else 0.20
    print("\n--- DCF assumptions: press Enter to accept the suggested default ---")
    years = int(ask_number("Forecast years (3-10)", 5, 3))
    years = min(years, 10)
    growth = ask_float("Annual revenue growth", 0.08, -0.50, 1.00)
    ebitda_margin = ask_float("EBITDA margin", margin if 0 < margin < 1 else 0.20, -0.50, 0.80)
    da_pct = ask_float("D&A as % of revenue", 0.04, 0, 0.50)
    capex_pct = ask_float("CapEx as % of revenue", 0.05, 0, 0.80)
    nwc_pct = ask_float("Change in NWC as % of incremental revenue", 0.02, -0.50, 0.50)
    tax = ask_float("Tax rate", 0.25, 0, 0.60)
    risk_free = ask_float("Risk-free rate", 0.04, 0, 0.30)
    beta = ask_number("Levered beta", 1.0, 0)
    erp = ask_float("Equity risk premium", 0.055, 0, 0.20)
    cost_debt = ask_float("Pre-tax cost of debt", 0.06, 0, 0.50)
    debt_weight = ask_float("Debt as % of total capital", 0.20, 0, 0.95)
    terminal_growth = ask_float("Terminal growth rate (must be below WACC)", 0.025, -0.05, 0.08)
    cash = ask_number("Latest cash and cash equivalents (same units as financials)", as_number(last.get("Cash & Cash Equivalents"), 0), 0)
    debt = ask_number("Latest total debt (same units as financials)", as_number(last.get("Total Debt"), 0), 0)
    if shares <= 0:
        shares = ask_number("Shares outstanding (in the same share units used for price calculation)", 1000, 0)
    if current_price <= 0:
        current_price = ask_number("Current share price (optional; enter 0 if unavailable)", 0, 0)
    return {
        "Company": company, "Forecast Years": years, "Revenue Growth": growth,
        "EBITDA Margin": ebitda_margin, "D&A % Revenue": da_pct,
        "CapEx % Revenue": capex_pct, "Change NWC % Incremental Revenue": nwc_pct,
        "Tax Rate": tax, "Risk-Free Rate": risk_free, "Beta": beta,
        "Equity Risk Premium": erp, "Pre-tax Cost of Debt": cost_debt,
        "Debt Weight": debt_weight, "Terminal Growth": terminal_growth,
        "Cash": cash, "Debt": debt, "Shares Outstanding": shares,
        "Current Share Price": current_price
    }

def style_title(ws, title: str, end_col: int):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_col)
    cell = ws.cell(1, 1, title)
    cell.fill = PatternFill("solid", fgColor=NAVY)
    cell.font = Font(name="Aptos Display", size=16, bold=True, color=WHITE)
    cell.alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 30

def style_header(row):
    for cell in row:
        cell.fill = PatternFill("solid", fgColor=DARK_BLUE)
        cell.font = Font(name="Aptos", bold=True, color=WHITE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=MED_BLUE)

def polish_sheet(ws):
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "B4"
    for row in ws.iter_rows():
        for cell in row:
            cell.font = cell.font.copy(name="Aptos")
            cell.alignment = cell.alignment.copy(vertical="center")
    for col in range(1, ws.max_column + 1):
        letter = get_column_letter(col)
        max_len = 0
        for cell in ws[letter]:
            if cell.value is not None:
                max_len = max(max_len, min(len(str(cell.value)), 36))
        ws.column_dimensions[letter].width = min(max(max_len + 3, 13), 30)
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_view.zoomScale = 90

def write_workbook(financials: pd.DataFrame, assumptions: dict, output_path: str, source_label: str):
    wb = Workbook()
    ws_fin = wb.active
    ws_fin.title = "Financials"
    ws_ass = wb.create_sheet("Assumptions")
    ws_fc = wb.create_sheet("Forecast")
    ws_dcf = wb.create_sheet("DCF & Sensitivity")
    ws_dash = wb.create_sheet("Charts & Dashboard")

    years = list(financials["Fiscal Year"].astype(int))
    style_title(ws_fin, f"{assumptions['Company']} | Historical Financials", len(years) + 2)
    ws_fin["A2"] = f"Source: {source_label} | Amounts and share count DISPLAYED in millions; underlying values unchanged. Enter uploaded financials in base units, not millions."
    ws_fin["A2"].font = Font(italic=True, color="666666", size=9)
    ws_fin.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(years)+2)
    ws_fin.append(["Metric"] + years + ["Data Quality / Notes"])
    style_header(ws_fin[3])
    items = LINE_ITEMS
    for item in items:
        row = [item]
        for year in years:
            match = financials.loc[financials["Fiscal Year"] == year]
            value = match.iloc[0].get(item) if not match.empty else None
            row.append(None if pd.isna(value) else value)
        notes = ""
        if item in ["EBITDA", "Gross Profit", "EBIT", "Pretax Income"] and all(v is None for v in row[1:-1]):
            notes = "Not available; can be derived or entered manually"
        elif item in ["Cash & Cash Equivalents", "Total Debt", "Shares Outstanding"] and all(v is None for v in row[1:-1]):
            notes = "Not available from source"
        row.append(notes)
        ws_fin.append(row)
    for row in ws_fin.iter_rows(min_row=4, min_col=2, max_col=len(years)+1):
        for cell in row:
            cell.number_format = MILLIONS_FMT
    for row in range(4, 4 + len(items)):
        ws_fin.cell(row, 1).font = Font(bold=True, color=NAVY)
        if ws_fin.cell(row, len(years)+2).value:
            ws_fin.cell(row, len(years)+2).fill = PatternFill("solid", fgColor=INPUT_YELLOW)
    ws_fin.auto_filter.ref = f"A3:{get_column_letter(len(years)+2)}{3+len(items)}"

    style_title(ws_ass, f"{assumptions['Company']} | User Assumptions", 4)
    ws_ass.append(["Assumption", "Value", "Type", "Guidance"])
    style_header(ws_ass[2])
    guidance = {
        "Revenue Growth": "Annual forecast revenue growth",
        "EBITDA Margin": "EBITDA divided by revenue",
        "D&A % Revenue": "Depreciation and amortization as a share of revenue",
        "CapEx % Revenue": "Capital expenditures as a share of revenue",
        "Change NWC % Incremental Revenue": "Change in net working capital as a share of incremental revenue",
        "Tax Rate": "Effective cash tax assumption",
        "Risk-Free Rate": "Risk-free rate used in CAPM",
        "Beta": "Levered equity beta",
        "Equity Risk Premium": "Expected market premium over risk-free rate",
        "Pre-tax Cost of Debt": "Estimated borrowing cost before tax",
        "Debt Weight": "Debt / (Debt + Equity)",
        "Terminal Growth": "Perpetual growth rate; must be below WACC",
        "Cash": "Latest cash balance, in same units as financials",
        "Debt": "Latest total debt, in same units as financials",
        "Shares Outstanding": "Enter the full share count, not millions; displayed in millions in the workbook.",
        "Current Share Price": "Used for upside/downside only"
    }
    input_rows = {}
    for key, value in assumptions.items():
        if key == "Company":
            continue
        row = ws_ass.max_row + 1
        input_rows[key] = row
        ws_ass.append([key, value, "User input", guidance.get(key, "Review and confirm")])
        ws_ass.cell(row, 2).fill = PatternFill("solid", fgColor=INPUT_YELLOW)
        ws_ass.cell(row, 2).font = Font(bold=True, color="7F6000")
        if isinstance(value, float) and any(s in key for s in ["Growth", "Margin", "Rate", "Premium", "Weight", "Cost", "Tax", "Revenue", "Debt Weight"]):
            ws_ass.cell(row, 2).number_format = "0.0%"
        elif key == "Beta":
            ws_ass.cell(row, 2).number_format = "0.00x"
        else:
            ws_ass.cell(row, 2).number_format = MILLIONS_FMT_2DP if key in ['Cash', 'Debt', 'Shares Outstanding'] else '#,##0.00;[Red](#,##0.00);-'
    ws_ass["A21"] = "WACC (calculated)"
    ws_ass["B21"] = f"=B{input_rows['Risk-Free Rate']}+B{input_rows['Beta']}*B{input_rows['Equity Risk Premium']}"
    ws_ass["C21"] = "Calculated"
    ws_ass["D21"] = "Cost of equity = risk-free rate + beta × equity risk premium"
    ws_ass["A22"] = "After-tax cost of debt"
    ws_ass["B22"] = f"=B{input_rows['Pre-tax Cost of Debt']}*(1-B{input_rows['Tax Rate']})"
    ws_ass["C22"] = "Calculated"
    ws_ass["D22"] = "Pre-tax cost of debt × (1 − tax rate)"
    ws_ass["A23"] = "WACC"
    ws_ass["B23"] = f"=B21*(1-B{input_rows['Debt Weight']})+B22*B{input_rows['Debt Weight']}"
    ws_ass["C23"] = "Calculated"
    ws_ass["D23"] = "Weighted cost of equity and after-tax cost of debt"
    for r in [21,22,23]:
        ws_ass.cell(r, 1).font = Font(bold=True, color=NAVY)
        ws_ass.cell(r, 2).fill = PatternFill("solid", fgColor=GREEN)
        ws_ass.cell(r, 2).number_format = "0.0%"

    forecast_years = [years[-1] + i for i in range(1, int(assumptions["Forecast Years"]) + 1)]
    style_title(ws_fc, f"{assumptions['Company']} | Forecast Income Statement & UFCF", len(forecast_years)+1)
    ws_fc.append(["Metric"] + forecast_years)
    style_header(ws_fc[2])
    forecast_metrics = [
        "Revenue", "Revenue Growth", "EBITDA Margin", "EBITDA", "D&A", "EBIT",
        "Cash Taxes on EBIT", "NOPAT", "Add: D&A", "Less: CapEx", "Less: Change in NWC", "Unlevered Free Cash Flow"
    ]
    fin_revenue_row = 4 + LINE_ITEMS.index("Revenue")
    fin_da_row = 4 + LINE_ITEMS.index("Depreciation & Amortization")
    row_map = {}
    for metric in forecast_metrics:
        row = ws_fc.max_row + 1
        row_map[metric] = row
        ws_fc.cell(row, 1, metric)
    rev_row = row_map["Revenue"]
    for col_idx in range(2, len(forecast_years)+2):
        col_letter = get_column_letter(col_idx)
        prev_col = get_column_letter(col_idx-1)
        hist_rev_col = get_column_letter(len(years)+1)
        ws_fc.cell(rev_row, col_idx, f"='Financials'!{hist_rev_col}{fin_revenue_row}" if col_idx == 2 else f"={prev_col}{rev_row}*(1+'Assumptions'!$B${input_rows['Revenue Growth']})")
        if col_idx == 2:
            ws_fc.cell(rev_row, col_idx, f"='Financials'!{hist_rev_col}{fin_revenue_row}*(1+'Assumptions'!$B${input_rows['Revenue Growth']})")
        ws_fc.cell(row_map["Revenue Growth"], col_idx, f"='Assumptions'!$B${input_rows['Revenue Growth']}")
        ws_fc.cell(row_map["EBITDA Margin"], col_idx, f"='Assumptions'!$B${input_rows['EBITDA Margin']}")
        ws_fc.cell(row_map["EBITDA"], col_idx, f"={col_letter}{rev_row}*{col_letter}{row_map['EBITDA Margin']}")
        ws_fc.cell(row_map["D&A"], col_idx, f"={col_letter}{rev_row}*'Assumptions'!$B${input_rows['D&A % Revenue']}")
        ws_fc.cell(row_map["EBIT"], col_idx, f"={col_letter}{row_map['EBITDA']}-{col_letter}{row_map['D&A']}")
        ws_fc.cell(row_map["Cash Taxes on EBIT"], col_idx, f"=MAX(0,{col_letter}{row_map['EBIT']}*'Assumptions'!$B${input_rows['Tax Rate']})")
        ws_fc.cell(row_map["NOPAT"], col_idx, f"={col_letter}{row_map['EBIT']}-{col_letter}{row_map['Cash Taxes on EBIT']}")
        ws_fc.cell(row_map["Add: D&A"], col_idx, f"={col_letter}{row_map['D&A']}")
        ws_fc.cell(row_map["Less: CapEx"], col_idx, f"={col_letter}{rev_row}*'Assumptions'!$B${input_rows['CapEx % Revenue']}")
        if col_idx == 2:
            prev_rev = f"'Financials'!{hist_rev_col}{fin_revenue_row}"
        else:
            prev_rev = f"{prev_col}{rev_row}"
        ws_fc.cell(row_map["Less: Change in NWC"], col_idx, f"=({col_letter}{rev_row}-{prev_rev})*'Assumptions'!$B${input_rows['Change NWC % Incremental Revenue']}")
        ws_fc.cell(row_map["Unlevered Free Cash Flow"], col_idx, f"={col_letter}{row_map['NOPAT']}+{col_letter}{row_map['Add: D&A']}-{col_letter}{row_map['Less: CapEx']}-{col_letter}{row_map['Less: Change in NWC']}")
    for row in range(3, ws_fc.max_row+1):
        ws_fc.cell(row, 1).font = Font(bold=(ws_fc.cell(row,1).value in ["Revenue","EBITDA","EBIT","NOPAT","Unlevered Free Cash Flow"]), color=NAVY)
    for row in range(3, ws_fc.max_row+1):
        for col_idx in range(2, len(forecast_years)+2):
            ws_fc.cell(row, col_idx).number_format = "0.0%" if ws_fc.cell(row,1).value in ["Revenue Growth","EBITDA Margin"] else MILLIONS_FMT
    ufcf_row = row_map["Unlevered Free Cash Flow"]

    style_title(ws_dcf, f"{assumptions['Company']} | DCF Valuation & Sensitivity", 8)
    ws_dcf.append(["Valuation Metric", "Value", "", "WACC / Terminal Growth Sensitivity", "", "", "", ""])
    style_header(ws_dcf[2])
    valuation_labels = [
        "WACC", "Terminal Growth", "PV of Forecast UFCF", "Terminal Value",
        "PV of Terminal Value", "Enterprise Value", "Add: Cash", "Less: Debt",
        "Equity Value", "Shares Outstanding", "Implied Share Value", "Current Share Price",
        "Upside / (Downside)"
    ]
    val_rows = {}
    for label in valuation_labels:
        r = ws_dcf.max_row + 1
        val_rows[label] = r
        ws_dcf.cell(r,1,label)
    wacc_ref = "'Assumptions'!$B$23"
    tg_ref = f"'Assumptions'!$B${input_rows['Terminal Growth']}"
    for i, col_idx in enumerate(range(2, len(forecast_years)+2), start=1):
        col_letter = get_column_letter(col_idx)
        ws_dcf.cell(val_rows["PV of Forecast UFCF"], 2, f"=SUM(" + ",".join([f"'{ 'Forecast' }'!{get_column_letter(j)}{ufcf_row}/(1+{wacc_ref})^{j-1}" for j in range(2, len(forecast_years)+2)]) + ")")
        break
    last_col = get_column_letter(len(forecast_years)+1)
    terminal_row_formula = f"='Forecast'!{last_col}{ufcf_row}*(1+{tg_ref})/({wacc_ref}-{tg_ref})"
    ws_dcf.cell(val_rows["WACC"], 2, f"={wacc_ref}")
    ws_dcf.cell(val_rows["Terminal Growth"], 2, f"={tg_ref}")
    ws_dcf.cell(val_rows["Terminal Value"], 2, f"={terminal_row_formula[1:]}")
    ws_dcf.cell(val_rows["PV of Terminal Value"], 2, f"=B{val_rows['Terminal Value']}/(1+B{val_rows['WACC']})^{len(forecast_years)}")
    ws_dcf.cell(val_rows["Enterprise Value"], 2, f"=B{val_rows['PV of Forecast UFCF']}+B{val_rows['PV of Terminal Value']}")
    ws_dcf.cell(val_rows["Add: Cash"], 2, f"='Assumptions'!$B${input_rows['Cash']}")
    ws_dcf.cell(val_rows["Less: Debt"], 2, f"='Assumptions'!$B${input_rows['Debt']}")
    ws_dcf.cell(val_rows["Equity Value"], 2, f"=B{val_rows['Enterprise Value']}+B{val_rows['Add: Cash']}-B{val_rows['Less: Debt']}")
    ws_dcf.cell(val_rows["Shares Outstanding"], 2, f"='Assumptions'!$B${input_rows['Shares Outstanding']}")
    ws_dcf.cell(val_rows["Implied Share Value"], 2, f"=IFERROR(B{val_rows['Equity Value']}/B{val_rows['Shares Outstanding']},0)")
    ws_dcf.cell(val_rows["Current Share Price"], 2, f"='Assumptions'!$B${input_rows['Current Share Price']}")
    ws_dcf.cell(val_rows["Upside / (Downside)"], 2, f"=IFERROR(B{val_rows['Implied Share Value']}/B{val_rows['Current Share Price']}-1,0)")
    for r in val_rows.values():
        ws_dcf.cell(r, 1).font = Font(bold=True, color=NAVY)
        label = ws_dcf.cell(r, 1).value
        if label in ["WACC", "Terminal Growth", "Upside / (Downside)"]:
            ws_dcf.cell(r, 2).number_format = "0.0%"
        elif label in ["Implied Share Value", "Current Share Price"]:
            ws_dcf.cell(r, 2).number_format = PER_SHARE_FMT
        else:
            ws_dcf.cell(r, 2).number_format = MILLIONS_FMT
    ws_dcf.cell(val_rows["Implied Share Value"],2).fill = PatternFill("solid", fgColor=GREEN)
    ws_dcf.cell(val_rows["Enterprise Value"],2).fill = PatternFill("solid", fgColor=LIGHT_BLUE)
    ws_dcf.cell(val_rows["Equity Value"],2).fill = PatternFill("solid", fgColor=LIGHT_BLUE)

    sens_start_row = 3
    sens_start_col = 4
    ws_dcf.cell(sens_start_row, sens_start_col, "WACC ↓ / g →")
    wacc_base = 0.08
    growth_base = as_number(assumptions["Terminal Growth"], 0.025)
    wacc_values = [max(0.01, as_number(0.08) + offset) for offset in [-0.02,-0.01,0,0.01,0.02]]
    growth_values = [growth_base + offset for offset in [-0.02,-0.01,0,0.01,0.02]]
    for j, growth in enumerate(growth_values, start=sens_start_col+1):
        ws_dcf.cell(sens_start_row, j, growth)
        ws_dcf.cell(sens_start_row, j).number_format = "0.0%"
    for i, wacc in enumerate(wacc_values, start=sens_start_row+1):
        ws_dcf.cell(i, sens_start_col, wacc)
        ws_dcf.cell(i, sens_start_col).number_format = "0.0%"
        for j in range(sens_start_col+1, sens_start_col+6):
            gcol = get_column_letter(j)
            forecast_pv = "+".join([f"'Forecast'!{get_column_letter(k)}{ufcf_row}/(1+$D{i})^{k-1}" for k in range(2, len(forecast_years)+2)])
            terminal_fcf = f"'Forecast'!{last_col}{ufcf_row}"
            ev_formula = f"=IFERROR(({forecast_pv}+({terminal_fcf}*(1+{gcol}${sens_start_row})/($D{i}-{gcol}${sens_start_row}))/(1+$D{i})^{len(forecast_years)}+'Assumptions'!$B${input_rows['Cash']}-'Assumptions'!$B${input_rows['Debt']})/'Assumptions'!$B${input_rows['Shares Outstanding']},0)"
            ws_dcf.cell(i,j,ev_formula)
            ws_dcf.cell(i,j).number_format = '#,##0.00;[Red](#,##0.00);-'
    for row in ws_dcf.iter_rows(min_row=sens_start_row, max_row=sens_start_row+5, min_col=sens_start_col, max_col=sens_start_col+5):
        for cell in row:
            cell.alignment = Alignment(horizontal="center")
            if cell.row == sens_start_row or cell.column == sens_start_col:
                cell.fill = PatternFill("solid", fgColor=DARK_BLUE)
                cell.font = Font(bold=True, color=WHITE)
    ws_dcf.conditional_formatting.add(f"E4:I8", ColorScaleRule(start_type="min", start_color="F8696B", mid_type="percentile", mid_value=50, mid_color="FFEB84", end_type="max", end_color="63BE7B"))
    ws_dcf["D10"] = "Sensitivity output: implied equity value per share"
    ws_dcf["D10"].font = Font(italic=True, color="666666", size=9)

    style_title(ws_dash, f"{assumptions['Company']} | DCF Dashboard", 12)
    ws_dash["A3"] = "Enterprise Value"
    ws_dash["D3"] = "Implied Share Value"
    ws_dash["G3"] = "WACC"
    ws_dash["J3"] = "Upside / (Downside)"
    cards = [
        ("A4", f"='DCF & Sensitivity'!B{val_rows['Enterprise Value']}"),
        ("D4", f"='DCF & Sensitivity'!B{val_rows['Implied Share Value']}"),
        ("G4", f"='DCF & Sensitivity'!B{val_rows['WACC']}"),
        ("J4", f"='DCF & Sensitivity'!B{val_rows['Upside / (Downside)']}")
    ]
    for label_cell, formula in cards:
        ws_dash[label_cell] = formula
        ws_dash[label_cell].fill = PatternFill("solid", fgColor=PALE_BLUE)
        ws_dash[label_cell].font = Font(size=16, bold=True, color=NAVY)
        ws_dash[label_cell].alignment = Alignment(horizontal="center", vertical="center")
        ws_dash[label_cell].number_format = "0.0%" if label_cell in ["G4", "J4"] else (PER_SHARE_FMT if label_cell == "D4" else MILLIONS_FMT)
    for cell in ["A3","D3","G3","J3"]:
        ws_dash[cell].fill = PatternFill("solid", fgColor=NAVY)
        ws_dash[cell].font = Font(bold=True, color=WHITE)
        ws_dash[cell].alignment = Alignment(horizontal="center")
    ws_dash.row_dimensions[4].height = 32
    ws_dash["A6"] = "Forecast Year"
    ws_dash["B6"] = "Revenue"
    ws_dash["C6"] = "EBITDA"
    ws_dash["D6"] = "UFCF"
    ws_dash["E6"] = "PV of UFCF"
    style_header(ws_dash[6][:5])
    for idx, year in enumerate(forecast_years, start=7):
        fc_col = get_column_letter(idx-5)
        ws_dash.cell(idx,1,year)
        ws_dash.cell(idx,2,f"='Forecast'!{fc_col}{row_map['Revenue']}")
        ws_dash.cell(idx,3,f"='Forecast'!{fc_col}{row_map['EBITDA']}")
        ws_dash.cell(idx,4,f"='Forecast'!{fc_col}{ufcf_row}")
        ws_dash.cell(idx,5,f"='Forecast'!{fc_col}{ufcf_row}/(1+'Assumptions'!$B$23)^{idx-6}")
        for col in range(2,6):
            ws_dash.cell(idx,col).number_format = MILLIONS_FMT
    bar = BarChart()
    bar.type = "col"
    bar.style = 10
    bar.title = "Revenue vs EBITDA"
    bar.y_axis.title = "Financial units (millions)"
    bar.y_axis.numFmt = MILLIONS_FMT
    bar.x_axis.title = "Forecast year"
    data = Reference(ws_dash, min_col=2, max_col=3, min_row=6, max_row=6+len(forecast_years))
    cats = Reference(ws_dash, min_col=1, min_row=7, max_row=6+len(forecast_years))
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)
    bar.height = 7
    bar.width = 13
    ws_dash.add_chart(bar, "G6")
    line = LineChart()
    line.title = "UFCF vs Present Value of UFCF"
    line.y_axis.title = "Financial units (millions)"
    line.y_axis.numFmt = MILLIONS_FMT
    line.x_axis.title = "Forecast year"
    data2 = Reference(ws_dash, min_col=4, max_col=5, min_row=6, max_row=6+len(forecast_years))
    line.add_data(data2, titles_from_data=True)
    line.set_categories(cats)
    line.height = 7
    line.width = 13
    ws_dash.add_chart(line, "G21")
    ws_dash["A14"] = "Model notes"
    ws_dash["A14"].font = Font(bold=True, color=WHITE)
    ws_dash["A14"].fill = PatternFill("solid", fgColor=NAVY)
    ws_dash.merge_cells("A14:E14")
    notes = [
        "This is an analytical model, not investment advice.",
        "Verify historical line items, units, share count, debt, cash and market price.",
        "Sensitivity table shows implied value per share under alternative WACC and terminal growth rates.",
        "Missing data are not a substitute for analyst review. Review all assumptions before use."
    ]
    for i, note in enumerate(notes, start=15):
        ws_dash.cell(i,1,"• " + note)
        ws_dash.merge_cells(start_row=i,start_column=1,end_row=i,end_column=5)
        ws_dash.cell(i,1).alignment = Alignment(wrap_text=True, vertical="top")
        ws_dash.row_dimensions[i].height = 24

    for ws in wb.worksheets:
        polish_sheet(ws)
        ws.sheet_properties.tabColor = NAVY
    ws_fin.column_dimensions["A"].width = 32
    ws_ass.column_dimensions["A"].width = 34
    ws_ass.column_dimensions["D"].width = 58
    ws_fc.column_dimensions["A"].width = 31
    ws_dcf.column_dimensions["A"].width = 31
    ws_dcf.column_dimensions["D"].width = 28
    ws_dash.column_dimensions["A"].width = 27
    ws_dash.column_dimensions["B"].width = 17
    ws_dash.column_dimensions["C"].width = 17
    ws_dash.column_dimensions["D"].width = 17
    ws_dash.column_dimensions["E"].width = 17
    ws_dash.freeze_panes = "B7"
    ws_dash.sheet_view.zoomScale = 85
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)

def main():
    parser = argparse.ArgumentParser(description="Build a generic DCF valuation workbook from scratch.")
    parser.add_argument("--ticker", help="Ticker symbol for yfinance, e.g. MSFT or RELIANCE.NS")
    parser.add_argument("--file", help="CSV/XLSX containing historical financials")
    parser.add_argument("--output", default="DCF_Valuation_Model.xlsx", help="Output Excel workbook path; financial amounts are displayed in millions")
    args = parser.parse_args()
    if bool(args.ticker) == bool(args.file):
        parser.error("Choose exactly one input: --ticker TICKER or --file historicals.xlsx")
    if args.ticker:
        financials, company, price, shares = load_yfinance_financials(args.ticker)
        source = f"yfinance ({args.ticker.upper()})"
    else:
        financials, company, price, shares = load_uploaded_financials(args.file)
        source = f"Uploaded file ({Path(args.file).name})"
    print(f"\nLoaded historical financials for {company}. Review the generated Financials sheet before relying on outputs.")
    assumptions = ask_assumptions(financials, company, price, shares)
    if assumptions["Terminal Growth"] >= 0.99:
        raise ValueError("Terminal growth is not valid.")
    write_workbook(financials, assumptions, args.output, source)
    print(f"\nWorkbook created: {Path(args.output).resolve()}")
    print("Open in Microsoft Excel or LibreOffice and allow formulas/charts to recalculate.")
    print("Important: confirm units, missing historical data, share count, debt/cash and all assumptions.")

if __name__ == "__main__":
    main()
