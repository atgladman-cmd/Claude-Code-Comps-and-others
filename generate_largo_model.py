#!/usr/bin/env python3
"""
Largo Resources (LGO) — Q2 2026 Quarterly Earnings Model
V2O5 price scenarios: US$5.00/lb  |  US$7.50/lb  |  US$10.00/lb

Sheets:
  1. Q2 2026 Model   – live market data header + three-scenario P&L + EV multiples
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
C_TEAL    = "1F7E79"
C_TEALLI  = "D6F0EF"

# Number formats
FMT_ACCT   = '_-"US$"* #,##0.0_-;[Red]_-"US$"* (#,##0.0)_-;_-"US$"* "-"??_-;_-@_-'
FMT_PCT    = '0.0%'
FMT_EPS    = '_-"US$"* #,##0.00_-;[Red]_-"US$"* (#,##0.00)_-;_-"US$"* "-"??_-;_-@_-'
FMT_SHARE  = '"US$"#,##0.000'
FMT_TONNE  = '#,##0'
FMT_MLBS   = '#,##0.00'
FMT_USD_T  = '"US$"#,##0'
FMT_PER_LB = '"US$"#,##0.00"/lb"'
FMT_MULT   = '#,##0.0"×"'
FMT_SHS_M  = '#,##0.0"M"'


# ─────────────────────────────────────────────────────────────────────────────
# STYLE HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def fnt(bold=False, sz=10, fc=C_BLACK, italic=False):
    return Font(name="Calibri", bold=bold, size=sz, color=fc, italic=italic)

def pfill(hex_col):
    return PatternFill("solid", fgColor=hex_col)

def aln(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def bdr(bc=C_BORDER):
    s = Side(style="thin", color=bc)
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

def wcell(ws, row, col, val=None, formula=None, **kwargs):
    c = ws.cell(row=row, column=col)
    if formula is not None:
        c.value = formula
    elif val is not None:
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

def sec_hdr(ws, row, c1, c2, text, bg=C_NAVY2):
    ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
    c = ws.cell(row=row, column=c1, value=text)
    sc(c, bold=True, sz=9, fc=C_WHITE, bg=bg, h="left", no_border=True)
    ws.row_dimensions[row].height = 14
    return c


def val_color(v, tol=0.0):
    if v is None:
        return C_LGREY, C_DGREY
    if v > tol:
        return C_GREEN, C_GREENDK
    if v < -tol:
        return C_RED, C_REDDK
    return C_YELLOW, C_YELLOWDK


# ─────────────────────────────────────────────────────────────────────────────
# MODEL INPUTS
# ─────────────────────────────────────────────────────────────────────────────
LB_PER_T    = 2204.6226
PRICES_LB   = [5.00, 7.50, 10.00]

# Q2 2026 cost assumptions (US$M) — anchored to Q1 2026 actuals (FactSet Final)
VOLUME_T    = 2400             # tonnes V2O5 sold Q2-2026 (est.)
ANN_PROD_T  = 9600             # annualised production (×4); used for EV/t
CASH_COGS   = 27.5             # cash operating costs ex-D&A  (Q1-26: $27.6M)
DA          = 7.0              # D&A  (Q1-26: $7.1M)
COGS_TOTAL  = CASH_COGS + DA   # = 34.5
SGA         = 4.5              # SG&A  (Q1-26: $4.5M)
NET_INT     = 2.5              # net interest expense (est. ~US$108M debt @ ~9.3%pa)
TAX_RATE    = 0.34             # Brazil CSLL/IRPJ; applied to positive EBT only

SHARES_M    = 88.31            # shares on issue (M) — Nov 2025 filing
IRESS_TICK  = "LGO"            # IRESS RTD ticker for LGO (NASDAQ, USD)

# Balance sheet — Q1 2026 (31 Mar 2026, FactSet Final, reported 13 May 2026)
TOTAL_DEBT  = 108.4            # US$M
NET_DEBT    = 96.8             # US$M  (= Total Debt – Cash of US$11.6M)
CASH_BS     = TOTAL_DEBT - NET_DEBT

# Breakeven prices (US$/lb) at VOLUME_T
VOL_MLBS    = VOLUME_T * LB_PER_T / 1e6
BE_EBITDA   = (CASH_COGS + SGA)            / VOL_MLBS
BE_EBIT     = (COGS_TOTAL + SGA)           / VOL_MLBS
BE_NI       = (COGS_TOTAL + SGA + NET_INT) / VOL_MLBS


def scen(price_lb):
    rev      = VOLUME_T * price_lb * LB_PER_T / 1e6
    gross    = rev - COGS_TOTAL
    ebit     = gross - SGA
    ebitda   = ebit + DA
    ebt      = ebit - NET_INT
    tax      = ebt * TAX_RATE if ebt > 0 else 0.0
    net      = ebt - tax
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
        ann_ebitda = ebitda * 4,
    )

SCENS = [scen(p) for p in PRICES_LB]

# Q1 2026 actuals (FactSet Final, reported 13 May 2026)
ACT = dict(
    rev        = 27.53,
    cogs       = 34.72,
    gross      = 27.53 - 34.72,
    gross_mgn  = (27.53 - 34.72) / 27.53,
    sga        = 4.54,
    ebit       = -4.59 - 7.13,
    ebit_mgn   = (-4.59 - 7.13) / 27.53,
    da         = 7.13,
    ebitda     = -4.59,
    ebitda_mgn = -4.59 / 27.53,
    tax        = 0.15,
    net        = -6.29,
    net_mgn    = -6.29 / 27.53,
    eps        = -6.29 / SHARES_M,
    ann_ebitda = -4.59 * 4,
)

# Historical actuals (Q1 2024 – Q1 2026, FactSet)
# (label, rev, cogs, sga, ebitda, da, net_inc, flag)
HIST = [
    ("Q1 2024", 42.19, 51.00, 6.46,  -6.55,  8.72, -12.97, ""),
    ("Q2 2024", 28.56, 38.41, 4.79,  -8.47,  6.06, -14.28, ""),
    ("Q3 2024", 29.91, 29.96, 8.57,  -3.04,  5.60,  -9.66, ""),
    ("Q4 2024", 24.27, 30.67, 1.35,   0.46,  7.70, -12.92, ""),
    ("Q1 2025", 28.24, 42.74, 5.02, -13.84,  5.68,  -9.00, ""),
    ("Q2 2025", 26.12, 30.31, 3.29,  -3.17,  4.48,  -5.67, ""),
    ("Q3 2025", 33.26, 34.65, 4.63,  -0.71,  5.36, -36.56, "†"),
    ("Q4 2025", 22.27, 26.04, 9.28,  -6.35,  6.51, -17.28, ""),
    ("Q1 2026", 27.53, 34.72, 4.54,  -4.59,  7.13,  -6.29, ""),
]


# ─────────────────────────────────────────────────────────────────────────────
# BUILD MODEL SHEET
# ─────────────────────────────────────────────────────────────────────────────
#  Col layout:
#  A=1  Label               width 36
#  B=2  Live / Q1-26 Actual width 15
#  C=3  separator           width  2
#  D=4  $5.00/lb Bear       width 15
#  E=5  $7.50/lb Base       width 15
#  F=6  $10.00/lb Bull      width 15

COL_LABEL = 1
COL_LIVE  = 2   # live mkt data / Q1-26 actual
COL_SEP   = 3
COL_S1    = 4   # Bear
COL_S2    = 5   # Base
COL_S3    = 6   # Bull
NCOLS     = 6
SCEN_COLS = [COL_S1, COL_S2, COL_S3]
SCEN_BGS  = [C_RED,  C_YELLOW, C_GREEN]
SCEN_TXT  = [C_REDDK, C_YELLOWDK, C_GREENDK]


def build_model(ws):

    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width =  2
    ws.column_dimensions["D"].width = 15
    ws.column_dimensions["E"].width = 15
    ws.column_dimensions["F"].width = 15

    R = 1   # running row pointer

    # ── Title ─────────────────────────────────────────────────────────────────
    ws.row_dimensions[R].height = 26
    mtitle(ws, R, 1, NCOLS,
           "LARGO RESOURCES (LGO)  —  Q2 2026 QUARTERLY EARNINGS MODEL",
           sz=13, ht=26)
    R += 1
    ws.row_dimensions[R].height = 14
    mtitle(ws, R, 1, NCOLS,
           "V₂O₅ price scenarios: US$5.00/lb  |  US$7.50/lb  |  US$10.00/lb  "
           " |  All figures US$ unless stated  |  Source: FactSet / IRESS RTD",
           bold=False, sz=9, bg=C_BLUE, ht=14)
    R += 1

    # ── spacer ────────────────────────────────────────────────────────────────
    ws.row_dimensions[R].height = 5
    for c in range(1, NCOLS + 1):
        sc(ws.cell(row=R, column=c), bg=C_WHITE, no_border=True)
    R += 1

    # ─────────────────────────────────────────────────────────────────────────
    # LIVE MARKET DATA BLOCK
    # Rows R=4 through R=13 (inclusive)
    # B-column cell addresses used for formula cross-references:
    #   ROW_PRICE   = R+1  (B = IRESS RTD price, US$/share)
    #   ROW_SHS     = R+2  (B = shares M)
    #   ROW_MKTCAP  = R+3  (B = =B_PRICE * B_SHS)
    #   ROW_ND      = R+4  (B = net debt US$M)
    #   ROW_EV      = R+5  (B = =B_MKTCAP + B_ND)
    #   ROW_EV_T    = R+6  (B = EV/t ann.)
    #   ROW_EV_LB   = R+7  (B = EV/lb ann.)
    # ─────────────────────────────────────────────────────────────────────────
    sec_hdr(ws, R, 1, NCOLS,
            f"  LIVE MARKET DATA  (share price via IRESS RTD — ticker: {IRESS_TICK})  "
            f"|  Net Debt: FactSet Q1 2026 Final (31 Mar 2026, reported 13 May 2026)",
            bg=C_TEAL)
    R += 1

    ROW_PRICE  = R
    ROW_SHS    = R + 1
    ROW_MKTCAP = R + 2
    ROW_ND     = R + 3
    ROW_EV     = R + 4
    ROW_EV_T   = R + 5
    ROW_EV_LB  = R + 6

    B_PRICE  = f"$B${ROW_PRICE}"
    B_SHS    = f"$B${ROW_SHS}"
    B_MKTCAP = f"$B${ROW_MKTCAP}"
    B_ND     = f"$B${ROW_ND}"
    B_EV     = f"$B${ROW_EV}"

    def live_row(label, val=None, formula=None, num_fmt=FMT_ACCT,
                 bold=False, note_lbl=None, row_ht=17):
        nonlocal R
        ws.row_dimensions[R].height = row_ht

        lc = ws.cell(row=R, column=COL_LABEL,
                     value=label + (f"  [{note_lbl}]" if note_lbl else ""))
        sc(lc, bold=bold, sz=10, fc=C_BLACK if not bold else C_NAVY,
           bg=C_TEALLI)

        vc = ws.cell(row=R, column=COL_LIVE)
        if formula:
            vc.value = formula
        elif val is not None:
            vc.value = val
        sc(vc, bold=bold, sz=11 if bold else 10, fc=C_NAVY, bg=C_TEALLI,
           h="right", num_fmt=num_fmt)

        # separator + scenario cols greyed out in this block
        sc(ws.cell(row=R, column=COL_SEP), bg=C_LGREY, no_border=True)
        for c in SCEN_COLS:
            sc(ws.cell(row=R, column=c), bg=C_LGREY, no_border=True)
        R += 1

    # LGO Last Price (live IRESS RTD)
    live_row("LGO  Last Price (US$/share) — NASDAQ",
             formula=f'=@IRESSRtd("Quote","13","0","{IRESS_TICK}")',
             num_fmt=FMT_SHARE, bold=True)

    live_row("Shares on Issue (millions)",
             val=SHARES_M, num_fmt=FMT_SHS_M)

    live_row(f"Market Capitalisation (US$M)  =  Price × {SHARES_M:.2f}M shares",
             formula=f"={B_PRICE}*{B_SHS}",
             bold=True)

    live_row("Net Debt (US$M)  —  31 Mar 2026  [FactSet FF_NET_DEBT, Q1 2026]",
             val=NET_DEBT)

    live_row("Enterprise Value (US$M)  =  Mkt Cap + Net Debt",
             formula=f"={B_MKTCAP}+{B_ND}",
             bold=True, row_ht=18)

    live_row(f"EV / Annualised V₂O₅ Production (US$/t)  [{ANN_PROD_T:,}t pa est.]",
             formula=f"={B_EV}*1000000/{ANN_PROD_T}",
             num_fmt=FMT_USD_T)

    live_row("EV / Annualised V₂O₅ Production (US$/lb)",
             formula=f"={B_EV}*1000000/{ANN_PROD_T}/{LB_PER_T}",
             num_fmt=FMT_PER_LB)

    # spacer
    ws.row_dimensions[R].height = 5
    for c in range(1, NCOLS + 1):
        sc(ws.cell(row=R, column=c), bg=C_WHITE, no_border=True)
    R += 1

    # ── Column headers ────────────────────────────────────────────────────────
    ws.row_dimensions[R].height = 36
    HDR_ROW = R
    hdr_items = [
        (COL_LABEL, "Line Item",                       "left",   C_NAVY,    C_WHITE),
        (COL_LIVE,  "Q1 2026\nActual",                 "center", C_NAVY,    C_WHITE),
        (COL_SEP,   "",                                "center", C_NAVY,    C_NAVY),
        (COL_S1,    "BEAR CASE\nUS$5.00/lb\nV₂O₅",    "center", "9C0006",  C_WHITE),
        (COL_S2,    "BASE CASE\nUS$7.50/lb\nV₂O₅",    "center", "7F6000",  C_WHITE),
        (COL_S3,    "BULL CASE\nUS$10.00/lb\nV₂O₅",   "center", "375623",  C_WHITE),
    ]
    for col, txt, align, bg, fg in hdr_items:
        wcell(ws, R, col, val=txt, bold=True, sz=9, fc=fg, bg=bg,
              h=align, v="center", wrap=True, thick_bot=True)
    DATA_START = R + 1
    R += 1

    # ── Inner helpers ─────────────────────────────────────────────────────────
    def blank_row():
        nonlocal R
        ws.row_dimensions[R].height = 5
        for c in range(1, NCOLS + 1):
            sc(ws.cell(row=R, column=c), bg=C_WHITE, no_border=True)
        R += 1

    def sec(label):
        nonlocal R
        sec_hdr(ws, R, 1, NCOLS, f"  {label}")
        R += 1

    # data_row: one full row with actual + 3 scenarios
    # row_type: "value" | "pct" | "mult"
    def data_row(label, act_val, s_vals,
                 bold=False, sz=10, indent=False,
                 row_type="value", use_color=False, tol=0.0, row_ht=16):
        nonlocal R
        ws.row_dimensions[R].height = row_ht

        lbl_bg = C_BGASMP if not indent else C_WHITE
        lc = ws.cell(row=R, column=COL_LABEL,
                     value=("    " + label if indent else label))
        sc(lc, bold=bold, sz=sz, fc=C_BLACK, bg=lbl_bg)

        # Actual column
        ac = ws.cell(row=R, column=COL_LIVE)
        if act_val is not None:
            ac.value = act_val
        nf = (FMT_PCT if row_type == "pct" else
              FMT_MULT if row_type == "mult" else FMT_ACCT)
        sc(ac, bold=bold, sz=sz, bg=C_LTBLUE, h="right", num_fmt=nf)

        # separator
        sc(ws.cell(row=R, column=COL_SEP), bg=C_LGREY, no_border=True)

        # scenario columns
        for ci, (col, sv) in enumerate(zip(SCEN_COLS, s_vals)):
            bg = SCEN_BGS[ci]
            fc = SCEN_TXT[ci]
            if use_color and sv is not None:
                bg, fc = val_color(sv, tol=tol)
            cell = ws.cell(row=R, column=col)
            if sv is not None:
                cell.value = sv
            sc(cell, bold=bold, sz=sz, fc=fc, bg=bg, h="right", num_fmt=nf)
        R += 1

    # ── Section: Operating Assumptions ───────────────────────────────────────
    sec("OPERATING ASSUMPTIONS")

    data_row("V₂O₅ Realised Price (US$/lb)",
             None, [s["price_lb"] for s in SCENS],
             bold=True, row_type="value")
    # patch fmt for price rows
    for c in SCEN_COLS:
        ws.cell(row=R-1, column=c).number_format = FMT_PER_LB

    data_row("V₂O₅ Realised Price (US$/t)",
             None, [s["price_t"] for s in SCENS])
    for c in SCEN_COLS:
        ws.cell(row=R-1, column=c).number_format = FMT_USD_T

    data_row("Sales Volume (t V₂O₅)  ⁽¹⁾",
             None, [s["vol_t"] for s in SCENS])
    for c in SCEN_COLS + [COL_LIVE]:
        ws.cell(row=R-1, column=c).number_format = FMT_TONNE

    data_row("Sales Volume (M lbs V₂O₅)",
             None, [s["vol_mlb"] for s in SCENS])
    for c in SCEN_COLS + [COL_LIVE]:
        ws.cell(row=R-1, column=c).number_format = FMT_MLBS

    blank_row()

    # ── Section: Income Statement ─────────────────────────────────────────────
    sec("INCOME STATEMENT  (US$M)")

    data_row("Revenue",
             ACT["rev"], [s["rev"] for s in SCENS], bold=True)

    data_row("Cost of Goods Sold (incl. D&A)  ⁽²⁾",
             ACT["cogs"], [s["cogs"] for s in SCENS], indent=True)

    data_row("Gross Profit / (Loss)",
             ACT["gross"], [s["gross"] for s in SCENS],
             bold=True, use_color=True)

    data_row("Gross Margin (%)",
             ACT["gross_mgn"], [s["gross_mgn"] for s in SCENS],
             row_type="pct", indent=True, use_color=True)

    blank_row()

    data_row("Selling, General & Administrative",
             ACT["sga"], [s["sga"] for s in SCENS], indent=True)

    data_row("EBIT",
             ACT["ebit"], [s["ebit"] for s in SCENS],
             bold=True, use_color=True)

    data_row("EBIT Margin (%)",
             ACT["ebit_mgn"], [s["ebit_mgn"] for s in SCENS],
             row_type="pct", indent=True, use_color=True)

    blank_row()

    data_row("Add: Depreciation, Depletion & Amortisation",
             ACT["da"], [s["da"] for s in SCENS], indent=True)

    data_row("EBITDA",
             ACT["ebitda"], [s["ebitda"] for s in SCENS],
             bold=True, sz=11, use_color=True, row_ht=18)

    data_row("EBITDA Margin (%)",
             ACT["ebitda_mgn"], [s["ebitda_mgn"] for s in SCENS],
             row_type="pct", indent=True, use_color=True)

    blank_row()

    data_row("Net Interest Expense  ⁽³⁾",
             None, [s["net_int"] for s in SCENS], indent=True)

    data_row("Earnings Before Tax (EBT)",
             None, [s["ebt"] for s in SCENS], use_color=True)

    data_row("Income Tax (34% on positive EBT)  ⁽⁴⁾",
             ACT["tax"], [s["tax"] for s in SCENS], indent=True)

    data_row("NET INCOME / (LOSS)",
             ACT["net"], [s["net"] for s in SCENS],
             bold=True, sz=11, use_color=True, row_ht=18)

    data_row("Net Margin (%)",
             ACT["net_mgn"], [s["net_mgn"] for s in SCENS],
             row_type="pct", indent=True, use_color=True)

    blank_row()

    # ── Section: Per Share ────────────────────────────────────────────────────
    sec("PER SHARE DATA")

    data_row("Shares on Issue (millions)",
             SHARES_M, [SHARES_M] * 3)
    for c in SCEN_COLS + [COL_LIVE]:
        ws.cell(row=R-1, column=c).number_format = FMT_SHS_M

    data_row("EPS (US$/share)",
             ACT["eps"], [s["eps"] for s in SCENS],
             bold=True, sz=11, use_color=True, row_ht=18)
    for c in SCEN_COLS + [COL_LIVE]:
        ws.cell(row=R-1, column=c).number_format = FMT_EPS

    blank_row()

    # ── Section: Enterprise Value & Multiples ─────────────────────────────────
    sec("ENTERPRISE VALUE & VALUATION MULTIPLES  (live EV from header above)")

    # Show the live EV in each scenario column — same value (EV is mkt-price-driven, not V2O5 price-driven)
    # Use a formula reference back to B_EV
    data_row("Enterprise Value — live (US$M)  ⁽⁵⁾",
             None, [None, None, None], bold=True)
    # override with formula in actual + scenario cols
    ws.cell(row=R-1, column=COL_LIVE).value  = f"={B_EV}"
    for c in SCEN_COLS:
        ws.cell(row=R-1, column=c).value = f"={B_EV}"
        sc(ws.cell(row=R-1, column=c), bold=True, sz=10,
           fc=C_NAVY, bg=C_TEALLI, h="right", num_fmt=FMT_ACCT)
    sc(ws.cell(row=R-1, column=COL_LIVE), bold=True, sz=10,
       fc=C_NAVY, bg=C_TEALLI, h="right", num_fmt=FMT_ACCT)

    data_row("Net Debt (US$M)",
             NET_DEBT, [NET_DEBT] * 3, indent=True)

    data_row("Implied Market Capitalisation (US$M)",
             None, [None, None, None], indent=True)
    for c in [COL_LIVE] + SCEN_COLS:
        ws.cell(row=R-1, column=c).value = f"={B_MKTCAP}"
        sc(ws.cell(row=R-1, column=c), sz=10,
           fc=C_NAVY if c == COL_LIVE else C_DGREY,
           bg=C_TEALLI if c == COL_LIVE else C_LGREY,
           h="right", num_fmt=FMT_ACCT)

    blank_row()

    # Annualised EBITDA per scenario (Q2 × 4 — indicative only)
    data_row("Annualised EBITDA  (Q2 × 4,  indicative)  (US$M)",
             ACT["ann_ebitda"], [s["ann_ebitda"] for s in SCENS],
             use_color=True)

    # EV/EBITDA multiples — live EV / annualised EBITDA
    # Use formula references; show n/m for negative EBITDA
    data_row("EV / Annualised EBITDA  ⁽⁶⁾",
             None, [None, None, None], bold=True, row_type="mult")

    evebitda_row = R - 1
    for ci, (c, s) in enumerate(zip(SCEN_COLS, SCENS)):
        ae = s["ann_ebitda"]
        if ae > 0:
            ws.cell(row=evebitda_row, column=c).value = f"={B_EV}/{ae:.4f}"
        else:
            ws.cell(row=evebitda_row, column=c).value = "n/m"
            sc(ws.cell(row=evebitda_row, column=c),
               bold=True, sz=10, fc=SCEN_TXT[ci], bg=SCEN_BGS[ci],
               h="right", no_border=False)
    sc(ws.cell(row=evebitda_row, column=COL_LIVE),
       bold=True, sz=10, bg=C_LTBLUE, h="right",
       num_fmt=FMT_MULT)
    ws.cell(row=evebitda_row, column=COL_LIVE).value = "n/m"

    blank_row()

    # EV per tonne of V2O5 — live (annualised)
    data_row("EV / Annualised V₂O₅ Production  (US$/t)  ⁽⁷⁾",
             None, [None, None, None])
    ev_t_row = R - 1
    for c in [COL_LIVE] + SCEN_COLS:
        ws.cell(row=ev_t_row, column=c).value = f"={B_EV}*1000000/{ANN_PROD_T}"
        sc(ws.cell(row=ev_t_row, column=c),
           fc=C_NAVY if c == COL_LIVE else C_DGREY,
           bg=C_TEALLI if c == COL_LIVE else C_LGREY,
           h="right", num_fmt=FMT_USD_T)

    data_row("EV / Annualised V₂O₅ Production  (US$/lb)  ⁽⁷⁾",
             None, [None, None, None])
    ev_lb_row = R - 1
    for c in [COL_LIVE] + SCEN_COLS:
        ws.cell(row=ev_lb_row, column=c).value = \
            f"={B_EV}*1000000/{ANN_PROD_T}/{LB_PER_T}"
        sc(ws.cell(row=ev_lb_row, column=c),
           fc=C_NAVY if c == COL_LIVE else C_DGREY,
           bg=C_TEALLI if c == COL_LIVE else C_LGREY,
           h="right", num_fmt=FMT_PER_LB)

    blank_row()

    # ── Section: Breakeven ────────────────────────────────────────────────────
    sec("BREAKEVEN ANALYSIS  (at {:,}t V₂O₅ sold)".format(VOLUME_T))

    for be_val, label in [
        (BE_EBITDA, "EBITDA Breakeven  (Revenue = Cash COGS + SG&A)"),
        (BE_EBIT,   "EBIT Breakeven     (Revenue = Total COGS + SG&A)"),
        (BE_NI,     "Net Income Breakeven  (EBT = 0;  COGS + SG&A + Interest)"),
    ]:
        ws.row_dimensions[R].height = 18
        lc = ws.cell(row=R, column=COL_LABEL, value=label)
        sc(lc, bold=True, sz=10, bg=C_BGASMP)
        ws.merge_cells(start_row=R, start_column=COL_LIVE,
                       end_row=R, end_column=COL_SEP)
        vc = ws.cell(row=R, column=COL_LIVE, value=be_val)
        sc(vc, bold=True, sz=11, fc=C_YELLOWDK, bg=C_YELLOW,
           h="right", num_fmt=FMT_PER_LB)
        for c in SCEN_COLS:
            sc(ws.cell(row=R, column=c), bg=C_YELLOW, no_border=True)
        R += 1

    blank_row()

    # ── Footnotes ─────────────────────────────────────────────────────────────
    mtitle(ws, R, 1, NCOLS, "  FOOTNOTES & MODEL NOTES",
           sz=9, bg=C_NAVY, h="left", ht=14)
    R += 1

    footnotes = [
        ("⁽¹⁾  Volume",
         f"{VOLUME_T:,}t V₂O₅ assumed for Q2 2026 (annualised ~{ANN_PROD_T:,}t pa). "
         "Based on Largo's recent run-rate (~2,200–2,500t/qtr) and FY2025 "
         "actual production of ~9,300t. Volume held constant across scenarios."),
        ("⁽²⁾  COGS",
         f"Includes D&A of US${DA:.1f}M/qtr. Cash COGS (ex-D&A) = US${CASH_COGS:.1f}M, "
         "anchored to Q1 2026 actuals (FactSet). Costs are BRL-denominated at mine "
         "level — a stronger/weaker BRL will shift reported USD costs."),
        ("⁽³⁾  Interest",
         f"Estimated US${NET_INT:.1f}M/qtr (US${TOTAL_DEBT:.0f}M total debt @ ~9.3%pa). "
         "Largo's debt is BRL-denominated; FX translation gains/losses on debt balances "
         "can materially affect reported Net Income and are NOT modelled here "
         "(see Q1 2026 actual where FX gains reduced net loss to US$6.3M)."),
        ("⁽⁴⁾  Tax",
         f"Brazil CSLL/IRPJ statutory rate {TAX_RATE:.0%}, applied to positive EBT only. "
         "Deferred tax assets on accumulated losses not assumed recognised."),
        ("⁽⁵⁾  Enterprise Value",
         f"EV = Live Market Cap (IRESS RTD price × {SHARES_M:.2f}M shares) + "
         f"Net Debt of US${NET_DEBT:.1f}M (FactSet FF_NET_DEBT, Q1 2026, 31 Mar 2026, "
         "reported 13 May 2026). EV moves in real-time with the share price. "
         f"IRESS ticker used: '{IRESS_TICK}' — verify exchange suffix (.NSQ for NASDAQ, "
         ".TSX for Toronto) if the RTD cell shows an error."),
        ("⁽⁶⁾  EV/EBITDA",
         "Annualised EBITDA = Q2 2026 modelled EBITDA × 4 (indicative; single-quarter "
         "annualisation). Shown as 'n/m' for negative EBITDA scenarios. "
         "Live EV from header divided by annualised EBITDA."),
        ("⁽⁷⁾  EV/V₂O₅",
         f"Live EV divided by annualised V₂O₅ production of {ANN_PROD_T:,}t pa "
         f"(= {VOLUME_T:,}t Q2 × 4). This metric is constant across all three "
         "price scenarios — it is driven solely by the live share price, not by "
         "the V₂O₅ commodity price assumption."),
    ]

    for label, text in footnotes:
        ws.row_dimensions[R].height = 14
        lc = ws.cell(row=R, column=COL_LABEL, value=label)
        sc(lc, bold=True, sz=8, fc=C_AMBERDK, bg=C_AMBER, h="left", no_border=True)
        ws.merge_cells(start_row=R, start_column=COL_LIVE,
                       end_row=R, end_column=NCOLS)
        tc = ws.cell(row=R, column=COL_LIVE, value=text)
        sc(tc, sz=8, italic=True, fc=C_AMBERDK, bg=C_AMBER,
           h="left", wrap=True, no_border=True)
        R += 1

    # freeze at first data row, name column
    ws.freeze_panes = f"B{DATA_START}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1
    ws.sheet_properties.tabColor = C_NAVY


# ─────────────────────────────────────────────────────────────────────────────
# BUILD HISTORICAL SHEET  (unchanged from v1)
# ─────────────────────────────────────────────────────────────────────────────
def build_history(ws):
    N = len(HIST)
    ws.column_dimensions["A"].width = 32
    for i in range(N):
        ws.column_dimensions[get_column_letter(2 + i)].width = 11

    R = 1
    ws.row_dimensions[R].height = 22
    mtitle(ws, R, 1, 1 + N,
           "LARGO RESOURCES (LGO)  —  QUARTERLY ACTUALS  |  Q1 2024 – Q1 2026",
           sz=12, ht=22)
    R += 1
    ws.row_dimensions[R].height = 13
    mtitle(ws, R, 1, 1 + N,
           "Source: FactSet Fundamentals API  "
           "(FF_SALES / FF_COGS / FF_SGA / FF_EBITDA_OPER / FF_DEP_EXP_CF / "
           "FF_INC_TAX / FF_NET_INC)  |  Currency: USD  |  All figures US$M",
           bold=False, sz=8, bg=C_BLUE, ht=13)
    R += 1
    ws.row_dimensions[R].height = 4
    R += 1

    # Quarter headers
    ws.row_dimensions[R].height = 28
    wcell(ws, R, 1, val="Line Item (US$M)", bold=True, sz=9,
          fc=C_WHITE, bg=C_NAVY, h="left", v="center", thick_bot=True)
    for i, h in enumerate(HIST):
        flag = h[7]
        wcell(ws, R, 2 + i,
              val=h[0] + (" " + flag if flag else ""),
              bold=True, sz=9, fc=C_WHITE, bg=C_NAVY,
              h="center", v="center", wrap=True, thick_bot=True)
    R += 1

    def h_sec(label):
        nonlocal R
        sec_hdr(ws, R, 1, 1 + N, f"  {label}")
        R += 1

    def h_row(label, vals, bold=False, indent=False,
              use_color=False, row_type="value", sz=9):
        nonlocal R
        ws.row_dimensions[R].height = 16
        lc = ws.cell(row=R, column=1,
                     value=("    " + label if indent else label))
        sc(lc, bold=bold, sz=sz, bg=C_BGASMP if not indent else C_WHITE)
        nf = (FMT_PCT  if row_type == "pct"  else
              FMT_EPS  if row_type == "eps"  else FMT_ACCT)
        for i, v in enumerate(vals):
            bg = C_LTBLUE if i % 2 == 0 else C_WHITE
            fc = C_BLACK
            if use_color and v is not None:
                bg, fc = val_color(v)
            cell = ws.cell(row=R, column=2 + i, value=v)
            sc(cell, bold=bold, sz=sz, fc=fc, bg=bg, h="right", num_fmt=nf)
        R += 1

    def h_blank():
        nonlocal R
        ws.row_dimensions[R].height = 5
        for c in range(1, 2 + N):
            sc(ws.cell(row=R, column=c), bg=C_WHITE, no_border=True)
        R += 1

    h_sec("INCOME STATEMENT  (US$M)")
    h_row("Revenue",                          [h[1] for h in HIST], bold=True)
    h_row("Cost of Goods Sold (incl. D&A)",   [h[2] for h in HIST], indent=True)

    gross = [h[1] - h[2] for h in HIST]
    h_row("Gross Profit / (Loss)", gross, bold=True, use_color=True)
    h_row("  Gross Margin (%)", [g/h[1] for g,h in zip(gross,HIST)],
          indent=True, use_color=True, row_type="pct")

    h_blank()
    h_row("SG&A",                             [h[3] for h in HIST], indent=True)

    ebit = [h[4] - h[5] for h in HIST]
    h_row("EBIT", ebit, bold=True, use_color=True)
    h_row("  EBIT Margin (%)", [e/h[1] for e,h in zip(ebit,HIST)],
          indent=True, use_color=True, row_type="pct")

    h_blank()
    h_row("Add: D&A",                         [h[5] for h in HIST], indent=True)
    h_row("EBITDA",                           [h[4] for h in HIST],
          bold=True, sz=10, use_color=True)
    h_row("  EBITDA Margin (%)", [h[4]/h[1] for h in HIST],
          indent=True, use_color=True, row_type="pct")

    h_blank()
    h_row("Net Income / (Loss)", [h[6] for h in HIST],
          bold=True, sz=10, use_color=True)
    h_row("  Net Margin (%)", [h[6]/h[1] for h in HIST],
          indent=True, use_color=True, row_type="pct")

    h_blank()
    h_sec("PER SHARE DATA")
    h_row("EPS (US$/share)", [h[6]/SHARES_M for h in HIST],
          bold=True, use_color=True, row_type="eps")

    h_blank()
    h_sec("BALANCE SHEET REFERENCE  (FactSet FF_NET_DEBT / FF_DEBT — most recent quarter)")
    h_row("Net Debt (US$M)  — Q1 2026", [None]*8 + [NET_DEBT])
    h_row("Total Debt (US$M) — Q1 2026", [None]*8 + [TOTAL_DEBT])
    h_row("Cash (US$M) — Q1 2026", [None]*8 + [round(CASH_BS,1)])

    h_blank()
    # footnote
    ws.merge_cells(start_row=R, start_column=1, end_row=R, end_column=1 + N)
    fn = ws.cell(row=R, column=1,
                 value="† Q3 2025: Net Loss of US$36.6M includes a large non-cash "
                       "impairment of the VCHARGE electrolyte business. Tax line reflects "
                       "a US$26.2M deferred tax benefit partly offsetting the write-down. "
                       "EBITDA of –US$0.7M was broadly in line with adjacent quarters.")
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
    print(f"  Breakeven EBITDA : US${BE_EBITDA:.2f}/lb")
    print(f"  Breakeven EBIT   : US${BE_EBIT:.2f}/lb")
    print(f"  Breakeven NI     : US${BE_NI:.2f}/lb")
    for s in SCENS:
        print(f"  ${s['price_lb']:.2f}/lb  Rev ${s['rev']:.1f}M  "
              f"EBITDA ${s['ebitda']:+.1f}M  Net ${s['net']:+.1f}M  "
              f"EPS ${s['eps']:+.2f}  AnnEBITDA ${s['ann_ebitda']:+.1f}M")


if __name__ == "__main__":
    main()
