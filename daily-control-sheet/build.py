"""
Build the daily control sheet workbook, or compare ours against one someone
else already made.

Nothing is cloned from a template — every cell, formula, and style is written
here, so the business rules live in code rather than in a file being passed
around.

Layout (CONTROL SHEET):
    row 2       headline totals and the deposit date
    rows 5-14   loan pulls        amount in G, label carries the name
    rows 17-36  trade tickets     amount in F, label carries the name
    rows 39-41  ACH & wires / remote deposit checks / LB 2999 checks
    row 42      totals and cross-foots
    row 44      ACH statement total and item count
    rows 46-55  minus PM ACH
    rows 57-66  plus wire
    rows 68-92  plus pulled ACH
    row 94      adjusted ACH & wires total
"""
from copy import copy
from datetime import date as _date

import openpyxl
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# --- layout ---------------------------------------------------------------
LOANS = (5, 14, "G")
TICKETS = (17, 36, "F")
PM_ACH = (46, 55, "C")
WIRES = (57, 66, "C")
PULLED = (68, 92, "C")

BLOCKS = {
    "loan_pulls": LOANS,
    "trade_tickets": TICKETS,
    "pm_ach": PM_ACH,
    "wires": WIRES,
    "pulled_ach": PULLED,
}

SCALARS = {
    "umb_check_count": "H2",
    "remote_deposit_count": "I2",
    "remote_deposit_amount": "E40",
    "lb2999_amount": "E41",
    "ach_statement_total": "C44",
    "ach_items": "E44",
}

# --- styling --------------------------------------------------------------
ACCT = r'_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)'
CURRENCY = r'"$"#,##0.00_);\("$"#,##0.00\)'
YELLOW = PatternFill("solid", fgColor="FFFFFF00")   # keyed in by hand
GREEN = PatternFill("solid", fgColor="FF92D050")    # worked out by the sheet
BLUE = PatternFill("solid", fgColor="FFDAEEF3")     # headline row shading
THIN = Border(bottom=Side(style="thin"))
MEDIUM = Border(bottom=Side(style="medium"))

def _box(style):
    s = Side(style=style)
    return Border(left=s, right=s, top=s, bottom=s)

THIN_BOX = _box("thin")
MEDIUM_BOX = _box("medium")

WIDTHS = {"A": 4.57, "B": 44.0, "C": 32.14, "D": 24.29, "E": 24.29,
          "F": 20.0, "G": 17.43, "H": 8.57, "J": 8.57, "K": 8.57}

# Taken from a finished sheet — the trade ticket rows are deep enough for a
# wrapped two-line name, the ACH blocks are tight.
HEIGHTS = {1: 55.5, 2: 26.25, 3: 20.25, 4: 15.75, 15: 18.0, 16: 18.0,
           37: 18.0, 38: 18.0, 39: 18.0, 40: 18.0, 41: 18.75, 42: 42.75,
           43: 21.75, 44: 17.25, 94: 15.75}
HEIGHTS.update({r: 18.0 for r in range(5, 15)})
HEIGHTS.update({r: 36.0 for r in range(17, 37)})
for _first, _last in ((46, 55), (57, 66), (68, 92)):
    HEIGHTS.update({r: 15.0 for r in range(_first, _last + 1)})

# Every populated cell above the totals row gets a full box; headers and
# totals get the heavy one. Below row 42 the sheet switches to a single
# outline down columns B and C, so it is handled separately.
MEDIUM_RANGES = ["D1:G1", "B2:G2", "B4:G4", "B42:G42"]
THIN_RANGES = ["H1:M1", "H2:M2", "B3:B3", "B15:D15", "B37:D37", "B39:E41"]
THIN_RANGES += [f"A{a}:G{b}" for a, b in ((5, 14), (17, 36))]

BLUE_CELLS = ["D1", "E1", "F1", "G1", "B2", "C2", "E2", "F2", "G2"]

# The lower half is framed as one tall outline rather than a grid: column B
# carries the left edge, column C the right, closed off top and bottom.
LOWER_FIRST, LOWER_LAST = 44, 94


def _put(ws, ref, value, *, size=14, bold=False, fill=None, fmt=None,
         border=None, align=None):
    cell = ws[ref]
    cell.value = value
    cell.font = Font(name="Arial", size=size, bold=bold)
    if fill:
        cell.fill = fill
    if fmt:
        cell.number_format = fmt
    if border:
        cell.border = border
    if align:
        cell.alignment = Alignment(horizontal=align)
    return cell


class Overflow(Exception):
    """More rows than the sheet has slots for."""


def _finish(ws):
    """
    Apply layout and borders once, after every value is in place.

    Doing it in one pass rather than per-cell means the whole grid is framed
    consistently, including cells left empty.
    """
    for col, width in WIDTHS.items():
        ws.column_dimensions[col].width = width
    for row, height in HEIGHTS.items():
        ws.row_dimensions[row].height = height

    for ranges, border in ((MEDIUM_RANGES, MEDIUM_BOX), (THIN_RANGES, THIN_BOX)):
        for span in ranges:
            for row in ws[span]:
                for cell in row:
                    cell.border = border

    for ref in BLUE_CELLS:
        ws[ref].fill = BLUE

    # One outline around the whole ACH reconciliation block, not a grid.
    thin, none = Side(style="thin"), Side()
    for row in range(LOWER_FIRST, LOWER_LAST + 1):
        top = thin if row == LOWER_FIRST else none
        bottom = thin if row == LOWER_LAST else none
        ws[f"B{row}"].border = Border(left=thin, right=none, top=top, bottom=bottom)
        ws[f"C{row}"].border = Border(left=none, right=thin, top=none, bottom=bottom)
    ws[f"E{LOWER_FIRST}"].border = Border(right=thin)

    # Row 1 headers centre on both axes and wrap.
    for cell in ws[1]:
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
    # Column B holds long descriptions, so it wraps too.
    for row in ws["B1:B94"]:
        for cell in row:
            keep = cell.alignment
            cell.alignment = Alignment(
                horizontal=keep.horizontal or ("center" if cell.row == 1 else None),
                vertical=keep.vertical or ("center" if cell.row == 1 else None),
                wrap_text=True)


def check_capacity(plan):
    problems = []
    for key, (first, last, _) in BLOCKS.items():
        slots = last - first + 1
        used = len(plan.get(key) or [])
        if used > slots:
            problems.append(f"{key}: {used} rows but only {slots} slots")
    if problems:
        raise Overflow("; ".join(problems))


# --- generation -----------------------------------------------------------
def generate(plan, path):
    """Write a complete workbook for `plan` to `path`."""
    check_capacity(plan)
    deposit = plan["date"]
    if isinstance(deposit, str):
        deposit = _date.fromisoformat(deposit)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CONTROL SHEET"

    # --- headline row ---
    for ref, text in (("D1", "DEPOSIT DATE"), ("E1", "PROCESS DATE"),
                      ("G1", "TOTAL # OF ITEMS")):
        _put(ws, ref, text, size=12, bold=True, border=MEDIUM, align="center")
    for ref, text in (("H1", "UMB Checks"), ("I1", "Remote Deposit Checks"),
                      ("J1", "ACHs & Wires"), ("K1", "Trade Tickets"),
                      ("M1", "Total Debits")):
        cell = _put(ws, ref, text, size=11, bold=True, align="center")
        cell.font = Font(name="Calibri", size=11, bold=True)

    _put(ws, "B2", "TOTAL DEPOSIT", size=12, bold=True, fmt=CURRENCY,
         border=MEDIUM, align="center")
    _put(ws, "C2", "=D42", size=18, bold=True, fmt=ACCT, border=MEDIUM, align="right")
    _put(ws, "D2", deposit, size=20, fill=YELLOW, fmt="mm/dd/yy",
         border=MEDIUM, align="center")
    _put(ws, "E2", "=WORKDAY(D2,1)", size=18, fmt="mm/dd/yy",
         border=MEDIUM, align="center")
    _put(ws, "G2", "=I2+J2+K2+H2", size=18, fmt="0", border=MEDIUM, align="center")
    _put(ws, "H2", plan.get("umb_check_count"), fill=YELLOW, align="center")
    _put(ws, "I2", plan.get("remote_deposit_count"), fill=YELLOW, fmt="0", align="center")
    _put(ws, "J2", "=E44-COUNTA(C46:C55)+COUNTA(C57:C66,C68:C92)",
         fill=GREEN, align="center")
    _put(ws, "K2", "=COUNTA(F17:F36)", fill=GREEN, align="center")
    _put(ws, "M2", "=COUNTA(F17:F36,C68:C92)", size=10, align="center")

    _put(ws, "B3", "CURRENT DATE LOCKBOX", size=11, bold=True)
    for ref, text in (("B4", "DESCRIPTION"), ("C4", "DEPOSIT DATE"),
                      ("D4", "TOTAL DEPOSIT"), ("E4", None), ("F4", None)):
        _put(ws, ref, text, size=11, bold=True, border=MEDIUM, align="center")

    # --- loan pulls and trade tickets ---
    def entry_block(first, last, amount_col, label, entries):
        for i, row in enumerate(range(first, last + 1)):
            item = entries[i] if i < len(entries) else None
            name = (item or {}).get("name", "") or ""
            _put(ws, f"A{row}", i + 1)
            _put(ws, f"B{row}", f"{label} ({name})", fmt="00000000",
                 border=THIN, align="left")
            _put(ws, f"C{row}", "=$D$2", fmt="mm/dd/yy", border=THIN, align="left")
            _put(ws, f"D{row}", f"={amount_col}{row}", fmt=ACCT,
                 border=THIN, align="left")
            for col in ("E", "F", "G"):
                if col != amount_col:
                    _put(ws, f"{col}{row}", None, fmt=ACCT, border=THIN, align="left")
            _put(ws, f"{amount_col}{row}", (item or {}).get("amount"),
                 fill=YELLOW, fmt=ACCT, border=THIN, align="left")

    entry_block(*LOANS[:2], LOANS[2], "CURRENT DATE LOAN PULL",
                plan.get("loan_pulls") or [])
    _put(ws, "B15", "TOTAL LOAN PULLS", bold=True, fmt="00000000",
         border=THIN, align="right")
    _put(ws, "C15", None, fmt="mm-dd-yy", border=THIN, align="left")
    _put(ws, "D15", "=SUM(D5:D14)", bold=True, fmt=ACCT, border=THIN, align="left")

    entry_block(*TICKETS[:2], TICKETS[2], "CURRENT DATE TRADE TICKET",
                plan.get("trade_tickets") or [])
    _put(ws, "B37", "TOTAL TRADE TICKETS", bold=True, fmt="00000000",
         border=THIN, align="right")
    _put(ws, "C37", None, fmt="mm-dd-yy", border=THIN, align="left")
    _put(ws, "D37", "=SUM(D17:D36)", bold=True, fmt=ACCT, border=THIN, align="left")

    # --- the three single-line entries ---
    singles = [
        (39, "ACH & Wires pushed & pulled", "=C94", False),
        (40, "Remote Deposit Checks", plan.get("remote_deposit_amount"), True),
        (41, "LB 2999 Checks", plan.get("lb2999_amount"), True),
    ]
    for row, label, value, keyed in singles:
        _put(ws, f"B{row}", label, fmt="00000000", border=THIN, align="left")
        _put(ws, f"C{row}", "=$D$2", fmt="mm/dd/yy", border=THIN, align="left")
        _put(ws, f"D{row}", f"=E{row}", bold=True, fmt=ACCT, border=THIN, align="left")
        _put(ws, f"E{row}", value, fill=YELLOW if keyed else None,
             fmt=ACCT, border=THIN, align="right")

    _put(ws, "B42", "TOTAL", size=11, bold=True, border=MEDIUM, align="left")
    _put(ws, "C42", None, size=11, fmt="mm-dd-yy", border=MEDIUM, align="left")
    for ref, formula in (("D42", "=SUM(D5:D14,D17:D36,D39:D41)"),
                         ("E42", "=SUM(E4:E41)"),
                         ("F42", "=SUM(F5:F41)"),
                         ("G42", "=SUM(G5:G41)")):
        _put(ws, ref, formula, bold=True, fmt=ACCT, border=MEDIUM, align="left")

    # --- ACH reconciliation ---
    _put(ws, "B44", "ACH Statement Total", size=11, bold=True)
    _put(ws, "C44", plan.get("ach_statement_total"), size=12, fill=YELLOW, fmt=ACCT)
    _put(ws, "D44", "Total ACH UMB items:", size=10, align="right")
    _put(ws, "E44", plan.get("ach_items"), size=12, fill=YELLOW)

    def amount_block(first, last, label, amounts):
        for i, row in enumerate(range(first, last + 1)):
            _put(ws, f"A{row}", i + 1, size=10)
            _put(ws, f"B{row}", label, size=10)
            _put(ws, f"C{row}", amounts[i] if i < len(amounts) else None,
                 size=12, fill=YELLOW, fmt=ACCT)

    amount_block(*PM_ACH[:2], "minus PM ACH", plan.get("pm_ach") or [])
    amount_block(*WIRES[:2], "plus wire", plan.get("wires") or [])
    amount_block(*PULLED[:2], "plus pulled ACH", plan.get("pulled_ach") or [])

    _put(ws, "B94", "Adjusted ACHs & Wires TOTAL", size=12, bold=True, border=THIN)
    _put(ws, "C94", "=C44-SUM(C46:C55)+SUM(C57:C66)+SUM(C68:C92)",
         size=12, bold=True, fmt=ACCT, border=THIN)

    # The sheet's own cross-foot checks. These cells carry no fill of their
    # own — a conditional format paints them green only while the total ties.
    # Differential styles colour from bgColor, not fgColor: setting fgColor
    # writes a rule that fires but shows nothing.
    good_fill = PatternFill(bgColor="FFC6EFCE", fill_type="solid")
    good_font = Font(color="FF006100")
    for ref, expected in (("D42", "SUM($D$15,$D$37,$D$39:$D$41)"),
                          ("E42", "SUM($D$39:$D$41)"),
                          ("F42", "$D$37"),
                          ("G42", "$D$15")):
        ws.conditional_formatting.add(ref, CellIsRule(
            operator="equal", formula=[expected],
            fill=good_fill, font=good_font))

    _finish(ws)

    # --- monthly rollup ---
    m = wb.create_sheet("Monthly numbers")
    for col, text in zip("ABCDEFG", ["Date", "Total  deposited", "UMB Checks",
                                     "Remote Deposit Checks", '"ACHs/Wires"',
                                     "TradeTickets", "ACH Debits"]):
        m[f"{col}1"] = text
    for col, formula in zip("ABCDEF", ["E2", "C2", "H2", "I2", "J2", "K2"]):
        m[f"{col}2"] = f"='CONTROL SHEET'!{formula}"
    m["G2"] = "='CONTROL SHEET'!M2-'CONTROL SHEET'!K2"
    m["A2"].number_format = "mm/dd/yy"
    for col in "BG":
        m[f"{col}2"].number_format = ACCT if col == "B" else "0"

    wb.save(path)
    return path


# --- reading and comparing ------------------------------------------------
def read_values(path):
    """Pull the input values back out of a workbook, ignoring blank rows."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["CONTROL SHEET"]

    out = {"date": ws["D2"].value}
    for key, ref in SCALARS.items():
        out[key] = ws[ref].value
    for key, (first, last, col) in BLOCKS.items():
        out[key] = [ws[f"{col}{r}"].value for r in range(first, last + 1)
                    if ws[f"{col}{r}"].value not in (None, "")]
    out["total_deposit"] = ws["D42"].value
    out["adjusted_ach"] = ws["C94"].value
    return out


def _same(a, b):
    """
    Exact comparison — an audit shouldn't wave anything through.

    Both sides come from the same decimal figures (parsed from the source
    files, or read back from Excel), so equal amounts land on the identical
    float and compare equal. A penny apart is a difference worth seeing.
    """
    if a is None or b is None:
        return a is None and b is None
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return float(a) == float(b)
    return a == b


def diff(existing_path, plan):
    """
    Compare a workbook someone else made against what we'd generate.

    Blocks are compared as ordered lists of the values actually present, not
    cell by cell, so a stray blank row between entries isn't reported as a
    difference — only the numbers matter.
    """
    theirs = read_values(existing_path)
    rows = []

    for key, ref in SCALARS.items():
        ours = plan.get(key)
        rows.append({
            "what": key, "cell": ref,
            "ours": ours, "theirs": theirs.get(key),
            "same": _same(ours, theirs.get(key)),
        })

    for key, (first, last, col) in BLOCKS.items():
        ours = [e["amount"] if isinstance(e, dict) else e
                for e in (plan.get(key) or [])]
        ours = [v for v in ours if v not in (None, "")]
        theirs_vals = [v for v in (theirs.get(key) or []) if v not in (None, "")]

        # Match value by value, not just on the total — two offsetting errors
        # would otherwise sum to the same figure and pass unnoticed.
        remaining = list(theirs_vals)
        only_ours = []
        for value in ours:
            hit = next((t for t in remaining if _same(value, t)), None)
            if hit is None:
                only_ours.append(value)
            else:
                remaining.remove(hit)

        rows.append({
            "what": key, "cell": f"{col}{first}:{col}{last}",
            "ours": ours, "theirs": theirs_vals,
            "only_ours": only_ours, "only_theirs": remaining,
            "ours_count": len(ours), "theirs_count": len(theirs_vals),
            "ours_total": round(sum(ours), 2),
            "theirs_total": round(sum(v for v in theirs_vals
                                      if isinstance(v, (int, float))), 2),
            "same": not only_ours and not remaining,
        })

    return {"rows": rows, "matches": all(r["same"] for r in rows),
            "their_total": theirs.get("total_deposit"),
            "their_adjusted_ach": theirs.get("adjusted_ach")}
