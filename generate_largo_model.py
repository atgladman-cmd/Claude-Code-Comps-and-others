#!/usr/bin/env python3
"""
Largo Resources (LGO) — Q2 2026 Quarterly Earnings Model
V2O5 price scenarios: US$5.00/lb  |  US$7.50/lb  |  US$10.00/lb

Sheets:
  1. Q2 2026 Model   – three-scenario P&L forecast
  2. Historical      – nine quarters of actuals (Q1 2024 – Q1 2026)
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────────────────────
# COLOUR PALETTE
# ─────────────────────────────────────────────────────────────────────────────
C_NAVY    = "1F4E79"
C_BLUE    = "2E75B6"
C_LTBLUE  = "D6E4F7"
C_WHITE   = "FFFFFF"
C_AMBER   = "FFF2CC"
C_AMBERDK = "7F6000"
C_LGREY   = "F2F2F2"
C_DGREY   = "595959"
C_BLACK   = "1A1A1A"
C_BORDER  = "9DC3E6"
C_NAVY2   = "17375E"
C_BGASMP  = "EBF3FB"
C_GREEN   = "E2EFDA"
C_YELLOW  = "FFEB9C"
C_RED     = "FCE4D6"
C_GREENDK = "375623"
C_REDDK   = "9C0006"
C_YELLOWDK= "7F6000"

# Number formats
FMT_ACCT  = '_-"$"* #,##0.0_-;[Red]_-"$"* (#,##0.0)_-;_-"$"* "-"??_-;_-@_-'
FMT_PCT   = '0.0%'
FMT_EPS   = '_-"$"* #,##0.00_-;[Red]_-"$"* (#,##0.00)_-;_-"$"* "-"??_-;_-@_-'
FMT_PRICE = '"US$"#,##0.00'
FMT_TONNE = '#,##0'
FMT_MLBS  = '#,##0.00'
FMT_USD_T = '"US$"#,##0'
FMT_PER_LB= '"US$"#,##0.00"/lb"'


# ─────────────────────────────────────────────────────────────────────────────
# STYLE HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def fnt(bold=False, sz=10, fc=C_BLACK, italic=False):
    return Font(name="Calibri", bold=bold, size=sz, color=fc, italic=italic)

def pfill(hex_col):
    return PatternFill("solid", fgColor=hex_col)

def aln(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def bdr(col=C_BORDER):
    s = Side(style="thin", color=col)
    return Border(left=s, right=s, top=s, bottom=s)

def bdr_none():
    return Border()

def bdr_thick_bot():
    t = Side(style="thin",   color=C_BORDER)
    b = Side(style="medium", color=C_NAVY)
    return Border(left=t, right=t, top=t, bottom=b)

def sc(cell, bold=False, sz=10, fc=C_BLACK, italic=False,
       bg=None, h="left", v="center", wrap=False,
       num_fmt=None, no_border=False, thick_bot=False):
    cell.font      = fnt(bold=bold, sz=sz, fc=fc, italic=italic)
    cell.alignment = aln(h=h, v=v, wrap=wrap)
    if bg:
        cell.fill = pfill(bg)
    if no_border:
        cell.border = bdr_none()
    elif thick_bot:
        cell.border = bdr_thick_bot()
    else:
        cell.border = bdr()
    if num_fmt:
        cell.number_format = num_fmt

def w(ws, row, col, val=None, **kwargs):
    c = ws.cell(row=row, column=col)
    if val is not None:
        c.value = val
    sc(c, **kwargs)
    return c

def mtitle(ws, row, c1, c2, text, bold=True, sz=12,
           fg=C_WHITE, bg=C_NAVY, h="center", ht=None):
    ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
    c = ws.cell(row=row, column=c1, value=text)
    sc(c, bold=bold, sz=sz, fc=fg, bg=bg, h=h, no_border=True)
    if ht:
        ws.row_dimensions[row].height = ht
    return c

def section_hdr(ws, row, c1, c2, text):
    ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
    c = ws.cell(row=row, column=c1, value=text)
    sc(c, bold=True, sz=9, fc=C_WHITE, bg=C_NAVY2, h="left", no_border=True)
    ws.row_dimensions[row].height = 14
    return c


# ─────────────────────────────────────────────────────────────────────────────
# MODEL INPUTS & CALCULATIONS
# ─────────────────────────────────────────────────────────────────────────────
LB_PER_T  = 2204.6226          # lb per metric tonne

PRICES_LB = [5.00, 7.50, 10.00]

# Q2 2026 cost assumptions (US$M) — anchored to Q1 2026 actuals (FactSet)
VOLUME_T   = 2400              # tonnes V2O5 sold (annualised ~9,600t; est.)
CASH_COGS  = 27.5              # cash operating costs ex D&A  (Q1-26: $27.6M)
DA         = 7.0               # D&A  (Q1-26: $7.1M)
COGS_TOTAL = CASH_COGS + DA    # = 34.5
SGA        = 4.5               # SG&A  (Q1-26: $4.5M)
NET_INT    = 2.5               # net interest expense (est. ~$108M debt @ ~9.3%)
TAX_RATE   = 0.34              # Brazilian CSLL/IRPJ (applied to +EBT only)

SHARES_M   = 88.31             # shares on issue (millions)

# Balance sheet (31 Mar 2026, FactSet)
TOTAL_DEBT = 108.4             # US$M
NET_DEBT   = 96.8              # US$M
CASH_BS    = TOTAL_DEBT - NET_DEBT   # ~11.6

# Breakeven prices (US$/lb)
VOL_MLBS   = VOLUME_T * LB_PER_T / 1e6
BE_EBITDA  = (CASH_COGS + SGA)       / VOL_MLBS
BE_EBIT    = (COGS_TOTAL + SGA)      / VOL_MLBS
BE_NI      = (COGS_TOTAL + SGA + NET_INT) / VOL_MLBS


def scen(price_lb):
    rev        = VOLUME_T * price_lb * LB_PER_T / 1e6
    gross      = rev - COGS_TOTAL
    ebit       = gross - SGA
    ebitda     = ebit + DA
    ebt        = ebit - NET_INT
    tax        = ebt * TAX_RATE if ebt > 0 else 0.0
    net        = ebt - tax
    return dict(
        price_lb   = price_lb,
        price_t    = price_lb * LB_PER_T,
        vol_t      = VOLUME_T,
        vol_mlb    = VOL_MLBS,
        rev        = rev,
        cogs       = COGS_TOTAL,
        gross      = gross,
        gross_mgn  = gross / rev,
        sga        = SGA,
        ebit       = ebit,
        ebit_mgn   = ebit / rev,
        da         = DA,
        ebitda     = ebitda,
        ebitda_mgn = ebitda / rev,
        net_int    = NET_INT,
        ebt        = ebt,
        tax        = tax,
        net        = net,
        net_mgn    = net / rev,
        eps        = net / SHARES_M,
    )

SCENS = [scen(p) for p in PRICES_LB]

# Q1 2026 actual (reference column — FactSet Final)
ACT = dict(
    label      = "Q1 2026\nActual",
    rev        = 27.53,
    cogs       = 34.72,
    gross      = 27.53 - 34.72,
    gross_mgn  = (27.53 - 34.72) / 27.53,
    sga        = 4.54,
    ebit       = -4.59 - 7.13,          # EBITDA – D&A
    ebit_mgn   = (-4.59 - 7.13) / 27.53,
    da         = 7.13,
    ebitda     = -4.59,
    ebitda_mgn = -4.59 / 27.53,
    net_int    = None,                   # see note
    ebt        = None,
    tax        = 0.15,
    net        = -6.29,
    net_mgn    = -6.29 / 27.53,
    eps        = -6.29 / SHARES_M,
)

# Historical actuals (Q1 2024 – Q1 2026, FactSet)
# (label, rev, cogs, sga, ebitda, da, net_inc, tax_note)
HIST = [
    ("Q1 2024", 42.19, 51.00, 6.46, -6.55,  8.72, -12.97, ""),
    ("Q2 2024", 28.56, 38.41, 4.79, -8.47,  6.06, -14.28, ""),
    ("Q3 2024", 29.91, 29.96, 8.57, -3.04,  5.60,  -9.66, ""),
    ("Q4 2024", 24.27, 30.67, 1.35,  0.46,  7.70, -12.92, ""),
    ("Q1 2025", 28.24, 42.74, 5.02,-13.84,  5.68,  -9.00, ""),
    ("Q2 2025", 26.12, 30.31, 3.29, -3.17,  4.48,  -5.67, ""),
    ("Q3 2025", 33.26, 34.65, 4.63, -0.71,  5.36, -36.56, "†"),
    ("Q4 2025", 22.27, 26.04, 9.28, -6.35,  6.51, -17.28, ""),
    ("Q1 2026", 27.53, 34.72, 4.54, -4.59,  7.13,  -6.29, ""),
]


# ─────────────────────────────────────────────────────────────────────────────
# BUILD MODEL SHEET
# ─────────────────────────────────────────────────────────────────────────────
# Column layout:
#  A=1  Label               width 34
#  B=2  Q1 2026 Actual      width 15
#  C=3  divider             width  2
#  D=4  $5.00/lb scenario   width 15
#  E=5  $7.50/lb scenario   width 15
#  F=6  $10.00/lb scenario  width 15

COL_LABEL = 1
COL_ACT   = 2
COL_SEP   = 3
COL_S1    = 4
COL_S2    = 5
COL_S3    = 6
NCOLS     = 6

SCEN_COLS = [COL_S1, COL_S2, COL_S3]
SCEN_BGS  = [C_RED, C_YELLOW, C_GREEN]
SCEN_TXT  = [C_REDDK, C_YELLOWDK, C_GREENDK]


def val_color(v, bg_pos=C_GREEN, bg_neg=C_RED, bg_zero=C_YELLOW, tol=0.0):
    if v is None:
        return C_LGREY, C_DGREY
    if v > tol:
        return bg_pos, C_GREENDK
    if v < -tol:
        return bg_neg, C_REDDK
    return bg_zero, C_YELLOWDK


def w_acct(ws, row, col, val, bg, **kwargs):
    return w(ws, row, col, val=val, bg=bg, num_fmt=FMT_ACCT, h="right", **kwargs)


def build_model(ws):
    # column widths
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width =  2
    ws.column_dimensions["D"].width = 15
    ws.column_dimensions["E"].width = 15
    ws.column_dimensions["F"].width = 15

    R = 1   # current row pointer

    # ── Title ─────────────────────────────────────────────────────────────────
    ws.row_dimensions[R].height = 26
    mtitle(ws, R, 1, NCOLS,
           "LARGO RESOURCES (LGO) — Q2 2026 QUARTERLY EARNINGS MODEL",
           sz=13, ht=26)
    R += 1

    ws.row_dimensions[R].height = 14
    mtitle(ws, R, 1, NCOLS,
           "V₂O₅ price sensitivity: US$5.00/lb  |  US$7.50/lb  |  US$10.00/lb  "
           " |  All figures US$M unless stated  |  Source: FactSet / company filings",
           bold=False, sz=9, bg=C_BLUE, ht=14)
    R += 1

    # ── Column headers ────────────────────────────────────────────────────────
    ws.row_dimensions[R].height = 4   # spacer
    R += 1
    ws.row_dimensions[R].height = 36

    hdr_items = [
        (COL_LABEL, "Line Item",         "left",    C_NAVY, C_WHITE),
        (COL_ACT,   "Q1 2026\nActual",   "center",  C_NAVY, C_WHITE),
        (COL_SEP,   "",                  "center",  C_NAVY, C_NAVY),
        (COL_S1,    "BEAR CASE\nUS$5.00/lb V₂O₅",  "center", "9C0006", C_WHITE),
        (COL_S2,    "BASE CASE\nUS$7.50/lb V₂O₅",  "center", "7F6000", C_WHITE),
        (COL_S3,    "BULL CASE\nUS$10.00/lb V₂O₅", "center", "375623", C_WHITE),
    ]
    for col, txt, align, bg, fg in hdr_items:
        w(ws, R, col, val=txt, bold=True, sz=9, fc=fg, bg=bg,
          h=align, v="center", wrap=True, thick_bot=True)
    R += 1

    # ── Helper: write one data row ────────────────────────────────────────────
    def data_row(label, act_val, s_vals, num_fmt=FMT_ACCT,
                 bold=False, sz=10, indent=False,
                 pct=False, tol=0.0, use_color=False,
                 row_ht=16):
        nonlocal R
        ws.row_dimensions[R].height = row_ht

        lbl_bg = C_BGASMP if not indent else C_WHITE
        lbl_col = C_BLACK

        # Label cell
        lc = ws.cell(row=R, column=COL_LABEL, value=("    " + label if indent else label))
        sc(lc, bold=bold, sz=sz, fc=lbl_col, bg=lbl_bg, h="left")

        # Actual column
        ac = ws.cell(row=R, column=COL_ACT)
        if act_val is not None:
            ac.value = act_val
        sc(ac, bold=bold, sz=sz, bg=C_LTBLUE, h="right",
           num_fmt=(FMT_PCT if pct else num_fmt))

        # Separator
        sc(ws.cell(row=R, column=COL_SEP), bg=C_LGREY, no_border=True)

        # Scenario columns
        for ci, (col, sv) in enumerate(zip(SCEN_COLS, s_vals)):
            sc_bg = SCEN_BGS[ci]
            sc_txt = SCEN_TXT[ci]
            if use_color:
                sc_bg, sc_txt = val_color(sv, tol=tol)
            cell = ws.cell(row=R, column=col)
            if sv is not None:
                cell.value = sv
            sc(cell, bold=bold, sz=sz, fc=sc_txt, bg=sc_bg, h="right",
               num_fmt=(FMT_PCT if pct else num_fmt))
        R += 1

    def blank_row():
        nonlocal R
        ws.row_dimensions[R].height = 5
        for col in range(1, NCOLS + 1):
            c = ws.cell(row=R, column=col)
            sc(c, bg=C_WHITE, no_border=True)
        R += 1

    def sec(label):
        nonlocal R
        section_hdr(ws, R, 1, NCOLS, f"  {label}")
        R += 1

    # ── Section 1: Operating Assumptions ──────────────────────────────────────
    sec("OPERATING ASSUMPTIONS")

    data_row("V₂O₅ Realised Price (US$/lb)",
             None,
             [s["price_lb"] for s in SCENS],
             num_fmt=FMT_PRICE, bold=True)

    data_row("V₂O₅ Realised Price (US$/t)",
             None,
             [s["price_t"] for s in SCENS],
             num_fmt=FMT_USD_T)

    data_row("Sales Volume (t V₂O₅)  ⁽¹⁾",
             None,
             [s["vol_t"] for s in SCENS],
             num_fmt=FMT_TONNE)

    data_row("Sales Volume (M lbs V₂O₅)",
             None,
             [s["vol_mlb"] for s in SCENS],
             num_fmt=FMT_MLBS)

    blank_row()

    # ── Section 2: Income Statement ───────────────────────────────────────────
    sec("INCOME STATEMENT (US$M)")

    data_row("Revenue",
             ACT["rev"],
             [s["rev"] for s in SCENS],
             bold=True)

    data_row("Cost of Goods Sold (incl. D&A)  ⁽²⁾",
             ACT["cogs"],
             [s["cogs"] for s in SCENS],
             indent=True)

    # Gross P&L row — colour by sign
    data_row("Gross Profit / (Loss)",
             ACT["gross"],
             [s["gross"] for s in SCENS],
             bold=True, use_color=True)

    data_row("Gross Margin (%)",
             ACT["gross_mgn"],
             [s["gross_mgn"] for s in SCENS],
             pct=True, use_color=True)

    blank_row()

    data_row("Selling, General & Administrative",
             ACT["sga"],
             [s["sga"] for s in SCENS],
             indent=True)

    data_row("EBIT",
             ACT["ebit"],
             [s["ebit"] for s in SCENS],
             bold=True, use_color=True)

    data_row("EBIT Margin (%)",
             ACT["ebit_mgn"],
             [s["ebit_mgn"] for s in SCENS],
             pct=True, use_color=True)

    blank_row()

    data_row("Add: Depreciation, Depletion & Amortisation",
             ACT["da"],
             [s["da"] for s in SCENS],
             indent=True)

    data_row("EBITDA",
             ACT["ebitda"],
             [s["ebitda"] for s in SCENS],
             bold=True, sz=11, use_color=True)

    data_row("EBITDA Margin (%)",
             ACT["ebitda_mgn"],
             [s["ebitda_mgn"] for s in SCENS],
             pct=True, use_color=True)

    blank_row()

    data_row("Net Interest Expense  ⁽³⁾",
             None,
             [s["net_int"] for s in SCENS],
             indent=True)

    data_row("Earnings Before Tax (EBT)",
             None,
             [s["ebt"] for s in SCENS],
             use_color=True)

    data_row("Income Tax  ⁽⁴⁾",
             ACT["tax"],
             [s["tax"] for s in SCENS],
             indent=True)

    data_row("NET INCOME / (LOSS)",
             ACT["net"],
             [s["net"] for s in SCENS],
             bold=True, sz=11, use_color=True)

    data_row("Net Margin (%)",
             ACT["net_mgn"],
             [s["net_mgn"] for s in SCENS],
             pct=True, use_color=True)

    blank_row()

    # ── Section 3: Per Share ──────────────────────────────────────────────────
    sec("PER SHARE DATA")

    data_row("Shares on Issue (millions)",
             SHARES_M,
             [SHARES_M] * 3,
             num_fmt=FMT_MLBS)

    data_row("EPS (US$/share)",
             ACT["eps"],
             [s["eps"] for s in SCENS],
             bold=True, sz=11, num_fmt=FMT_EPS, use_color=True)

    blank_row()

    # ── Section 4: Balance Sheet Reference ───────────────────────────────────
    sec("BALANCE SHEET REFERENCE  (31 Mar 2026 — FactSet)")

    data_row("Total Debt (US$M)",
             TOTAL_DEBT,
             [TOTAL_DEBT] * 3)

    data_row("Cash & Equivalents (US$M)",
             CASH_BS,
             [CASH_BS] * 3)

    data_row("Net Debt (US$M)",
             NET_DEBT,
             [NET_DEBT] * 3,
             bold=True)

    blank_row()

    # ── Section 5: Breakeven Analysis ────────────────────────────────────────
    sec("BREAKEVEN ANALYSIS  (at 2,400t V₂O₅ sold)")

    ws.row_dimensions[R].height = 18
    lc = ws.cell(row=R, column=COL_LABEL,
                 value="EBITDA Breakeven  (Revenue = Cash COGS + SG&A)")
    sc(lc, bold=True, sz=10, bg=C_BGASMP)
    ws.merge_cells(start_row=R, start_column=COL_ACT, end_row=R, end_column=COL_SEP)
    ac = ws.cell(row=R, column=COL_ACT, value=BE_EBITDA)
    sc(ac, bold=True, sz=11, fc=C_YELLOWDK, bg=C_YELLOW, h="right",
       num_fmt=FMT_PER_LB)
    for col in [COL_S1, COL_S2, COL_S3]:
        sc(ws.cell(row=R, column=col), bg=C_YELLOW, no_border=True)
    R += 1

    ws.row_dimensions[R].height = 18
    lc = ws.cell(row=R, column=COL_LABEL,
                 value="EBIT Breakeven  (Revenue = Total COGS + SG&A)")
    sc(lc, bold=True, sz=10, bg=C_BGASMP)
    ws.merge_cells(start_row=R, start_column=COL_ACT, end_row=R, end_column=COL_SEP)
    ac = ws.cell(row=R, column=COL_ACT, value=BE_EBIT)
    sc(ac, bold=True, sz=11, fc=C_YELLOWDK, bg=C_YELLOW, h="right",
       num_fmt=FMT_PER_LB)
    for col in [COL_S1, COL_S2, COL_S3]:
        sc(ws.cell(row=R, column=col), bg=C_YELLOW, no_border=True)
    R += 1

    ws.row_dimensions[R].height = 18
    lc = ws.cell(row=R, column=COL_LABEL,
                 value="Net Income Breakeven  (EBT = 0;  Revenue = COGS + SG&A + Interest)")
    sc(lc, bold=True, sz=10, bg=C_BGASMP)
    ws.merge_cells(start_row=R, start_column=COL_ACT, end_row=R, end_column=COL_SEP)
    ac = ws.cell(row=R, column=COL_ACT, value=BE_NI)
    sc(ac, bold=True, sz=11, fc=C_YELLOWDK, bg=C_YELLOW, h="right",
       num_fmt=FMT_PER_LB)
    for col in [COL_S1, COL_S2, COL_S3]:
        sc(ws.cell(row=R, column=col), bg=C_YELLOW, no_border=True)
    R += 1

    blank_row()

    # ── Footnotes ─────────────────────────────────────────────────────────────
    fn_data = [
        ("⁽¹⁾  Sales Volume",
         "2,400t V₂O₅ assumed for Q2 2026 (annualised ~9,600 tpa). "
         "Based on Largo's recent quarterly run-rate (~2,200–2,500t/qtr) and "
         "FY2025 actual production of ~9,300t. Volume held constant across scenarios — "
         "only the realised price changes."),
        ("⁽²⁾  COGS",
         f"Includes D&A of US${DA:.1f}M/qtr. Cash COGS (ex-D&A) = US${CASH_COGS:.1f}M, "
         "calibrated to Q1 2026 actuals (FactSet). COGS reflects Largo's "
         "Maracás Menchen mine (Brazil) operating costs, royalties, and on-site processing. "
         "Costs denominated in BRL — a stronger/weaker BRL vs. USD will affect reported USD costs."),
        ("⁽³⁾  Net Interest",
         f"Estimated US${NET_INT:.1f}M/qtr based on total debt of US${TOTAL_DEBT:.0f}M "
         "(FactSet, 31 Mar 2026) at an implied ~9.3% annual rate. "
         "Largo's actual debt is primarily BRL-denominated; FX movements on debt "
         "balances (translation gains/losses) can materially affect reported EBT and "
         "Net Income but are excluded from this model."),
        ("⁽⁴⁾  Tax",
         f"Brazilian CSLL/IRPJ statutory rate of {TAX_RATE:.0%} applied to "
         "positive EBT only. Deferred tax assets on accumulated losses are not "
         "assumed to be recognised. Q1 2026 actual tax of US$0.15M per FactSet."),
        ("Q1 2026 Actual",
         "Reported by Largo Resources on 13 May 2026 (Q1 2026 results). "
         "Net Income diverges from the modelled EBT path due to FX translation "
         "gains/losses on BRL-denominated debt (not modelled here). "
         "Revenue implies ~US$5.50/lb realised V₂O₅ price on ~2,300t sold."),
        ("Data Sources",
         "Income statement metrics (FF_SALES, FF_COGS, FF_SGA, FF_EBITDA_OPER, "
         "FF_DEP_EXP_CF, FF_INC_TAX, FF_NET_INC) and balance sheet metrics "
         "(FF_NET_DEBT, FF_DEBT) sourced from FactSet Fundamentals API, "
         f"currency USD, as at {__import__('datetime').date.today()}."),
    ]

    ws.row_dimensions[R].height = 4
    for col in range(1, NCOLS + 1):
        sc(ws.cell(row=R, column=col), bg=C_WHITE, no_border=True)
    R += 1

    mtitle(ws, R, 1, NCOLS, "  FOOTNOTES & MODEL NOTES",
           sz=9, bg=C_NAVY, h="left", ht=14)
    R += 1

    for label, text in fn_data:
        ws.row_dimensions[R].height = 14
        lc = ws.cell(row=R, column=COL_LABEL, value=label)
        sc(lc, bold=True, sz=8, fc=C_AMBERDK, bg=C_AMBER, h="left", no_border=True)
        ws.merge_cells(start_row=R, start_column=COL_ACT, end_row=R, end_column=NCOLS)
        tc = ws.cell(row=R, column=COL_ACT, value=text)
        sc(tc, sz=8, italic=True, fc=C_AMBERDK, bg=C_AMBER, h="left",
           wrap=True, no_border=True)
        R += 1

    # ── Freeze & print ────────────────────────────────────────────────────────
    ws.freeze_panes = "B5"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1
    ws.sheet_properties.tabColor = C_NAVY


# ─────────────────────────────────────────────────────────────────────────────
# BUILD HISTORICAL SHEET
# ─────────────────────────────────────────────────────────────────────────────
def build_history(ws):
    N = len(HIST)

    # Column widths: A=label, B-J=quarters
    ws.column_dimensions["A"].width = 32
    for i in range(N):
        ws.column_dimensions[get_column_letter(2 + i)].width = 11

    R = 1

    ws.row_dimensions[R].height = 22
    mtitle(ws, R, 1, 1 + N,
           "LARGO RESOURCES (LGO) — QUARTERLY ACTUALS  |  Q1 2024 – Q1 2026",
           sz=12, ht=22)
    R += 1

    ws.row_dimensions[R].height = 13
    mtitle(ws, R, 1, 1 + N,
           "Source: FactSet Fundamentals API (FF_SALES / FF_COGS / FF_SGA / "
           "FF_EBITDA_OPER / FF_DEP_EXP_CF / FF_INC_TAX / FF_NET_INC)  |  "
           "Currency: USD  |  All figures US$M",
           bold=False, sz=8, bg=C_BLUE, ht=13)
    R += 1

    ws.row_dimensions[R].height = 4  # spacer
    R += 1

    # Quarter header row
    ws.row_dimensions[R].height = 28
    w(ws, R, 1, val="Line Item (US$M)", bold=True, sz=9,
      fc=C_WHITE, bg=C_NAVY, h="left", v="center", thick_bot=True)
    for i, (qtr, *_) in enumerate(HIST):
        flag = HIST[i][7]
        w(ws, R, 2 + i, val=qtr + (" " + flag if flag else ""),
          bold=True, sz=9, fc=C_WHITE, bg=C_NAVY,
          h="center", v="center", wrap=True, thick_bot=True)
    R += 1

    def hist_section(label):
        nonlocal R
        ws.merge_cells(start_row=R, start_column=1, end_row=R, end_column=1 + N)
        c = ws.cell(row=R, column=1, value=f"  {label}")
        sc(c, bold=True, sz=9, fc=C_WHITE, bg=C_NAVY2, no_border=True)
        ws.row_dimensions[R].height = 14
        R += 1

    def hist_row(label, vals, num_fmt=FMT_ACCT, bold=False,
                 indent=False, use_color=False, pct=False, sz=9):
        nonlocal R
        ws.row_dimensions[R].height = 16
        prefix = "    " if indent else ""
        lc = ws.cell(row=R, column=1, value=prefix + label)
        sc(lc, bold=bold, sz=9, bg=C_BGASMP if not indent else C_WHITE)
        for i, v in enumerate(vals):
            col = 2 + i
            bg = C_LTBLUE if i % 2 == 0 else C_WHITE
            txt = C_BLACK
            if use_color and v is not None:
                bg, txt = val_color(v)
            cell = ws.cell(row=R, column=col, value=v)
            sc(cell, bold=bold, sz=9, fc=txt, bg=bg, h="right",
               num_fmt=(FMT_PCT if pct else num_fmt))
        R += 1

    def hblank():
        nonlocal R
        ws.row_dimensions[R].height = 5
        for col in range(1, 2 + N):
            sc(ws.cell(row=R, column=col), bg=C_WHITE, no_border=True)
        R += 1

    # ── Revenue ───────────────────────────────────────────────────────────────
    hist_section("INCOME STATEMENT (US$M)")
    hist_row("Revenue", [h[1] for h in HIST], bold=True)
    hist_row("Cost of Goods Sold (incl. D&A)", [h[2] for h in HIST], indent=True)

    gross_vals = [h[1] - h[2] for h in HIST]
    hist_row("Gross Profit / (Loss)", gross_vals, bold=True, use_color=True)

    gross_mgn = [(h[1] - h[2]) / h[1] for h in HIST]
    hist_row("  Gross Margin (%)", gross_mgn, pct=True, indent=True, use_color=True)

    hblank()

    hist_row("Selling, General & Administrative", [h[3] for h in HIST], indent=True)

    ebit_vals = [h[4] - h[5] for h in HIST]
    hist_row("EBIT", ebit_vals, bold=True, use_color=True)

    ebit_mgn = [(h[4] - h[5]) / h[1] for h in HIST]
    hist_row("  EBIT Margin (%)", ebit_mgn, pct=True, indent=True, use_color=True)

    hblank()

    hist_row("Add: D&A", [h[5] for h in HIST], indent=True)
    hist_row("EBITDA", [h[4] for h in HIST], bold=True, sz=10, use_color=True)

    ebitda_mgn = [h[4] / h[1] for h in HIST]
    hist_row("  EBITDA Margin (%)", ebitda_mgn, pct=True, indent=True, use_color=True)

    hblank()

    hist_row("Income Tax (expense) / benefit", [h[7] for h in HIST], indent=True)
    hist_row("Net Income / (Loss)", [h[6] for h in HIST],
             bold=True, sz=10, use_color=True)

    net_mgn = [h[6] / h[1] for h in HIST]
    hist_row("  Net Margin (%)", net_mgn, pct=True, indent=True, use_color=True)

    hblank()

    hist_section("PER SHARE DATA")
    eps_vals = [h[6] / SHARES_M for h in HIST]
    hist_row("EPS (US$/share)", eps_vals, bold=True, num_fmt=FMT_EPS, use_color=True)

    hblank()

    # ── Footnote ──────────────────────────────────────────────────────────────
    ws.merge_cells(start_row=R, start_column=1, end_row=R, end_column=1 + N)
    fn = ws.cell(row=R, column=1,
                 value="† Q3 2025: Net Loss of US$36.6M reflects a large non-cash impairment "
                       "(VCHARGE electrolyte business). Tax line shows a US$26.2M deferred "
                       "tax benefit, partially offsetting the write-down. "
                       "EBITDA of -US$0.7M was broadly in line with surrounding quarters.")
    sc(fn, sz=8, italic=True, fc=C_AMBERDK, bg=C_AMBER,
       h="left", wrap=True, no_border=True)
    ws.row_dimensions[R].height = 28

    ws.freeze_panes = "B5"
    ws.page_setup.orientation  = "landscape"
    ws.page_setup.fitToPage    = True
    ws.page_setup.fitToWidth   = 1
    ws.sheet_properties.tabColor = C_BLUE


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────
def main():
    wb = openpyxl.Workbook()

    ws_model = wb.active
    ws_model.title = "Q2 2026 Model"
    ws_model.sheet_properties.tabColor = C_NAVY

    ws_hist = wb.create_sheet("Historical")
    ws_hist.sheet_properties.tabColor = C_BLUE

    build_model(ws_model)
    build_history(ws_hist)

    out = "largo_v2o5_model.xlsx"
    wb.save(out)
    print(f"Saved: {out}")
    print(f"  Breakeven EBITDA : US${BE_EBITDA:.2f}/lb V2O5")
    print(f"  Breakeven EBIT   : US${BE_EBIT:.2f}/lb V2O5")
    print(f"  Breakeven NI     : US${BE_NI:.2f}/lb V2O5")
    for s in SCENS:
        print(f"  ${s['price_lb']:.2f}/lb → Revenue ${s['rev']:.1f}M  "
              f"EBITDA ${s['ebitda']:+.1f}M  Net ${s['net']:+.1f}M  "
              f"EPS ${s['eps']:+.2f}")


if __name__ == "__main__":
    main()
