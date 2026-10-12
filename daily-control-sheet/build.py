"""
Build the daily control sheet workbook, or compare ours against one someone
else already made.

Nothing is cloned from a template — every cell, formula, and style is written
here, so the business rules live in code rather than in a file being passed
around.

The layout is computed rather than fixed. Each block keeps its usual size
unless the day needs more rows, in which case it grows and everything below
shifts down. Formulas are built from those positions, so they stay correct at
any size:

    row 2                headline totals and the deposit date
    loan pulls           amount in G, label carries the name   (>= 10 rows)
    trade tickets        amount in F, label carries the name   (>= 20 rows)
    ACH & wires / remote deposit checks / LB 2999 checks
    TOTAL and cross-foots
    ACH statement total and item count
    minus PM ACH                                               (>= 10 rows)
    plus wire                                                  (>= 10 rows)
    plus pulled ACH                                            (>= 25 rows)
    adjusted ACH & wires total
"""
import re
from datetime import date as _date

import openpyxl
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# --- layout ---------------------------------------------------------------
SIZES = {
    "loan_pulls": 10,
    "trade_tickets": 20,
    "pm_ach": 10,
    "wires": 10,
    "pulled_ach": 25,
}
AMOUNT_COL = {
    "loan_pulls": "G", "trade_tickets": "F",
    "pm_ach": "C", "wires": "C", "pulled_ach": "C",
}
LABELS = {
    "loan_pulls": "CURRENT DATE LOAN PULL",
    "trade_tickets": "CURRENT DATE TRADE TICKET",
    "pm_ach": "minus PM ACH",
    "wires": "plus wire",
    "pulled_ach": "plus pulled ACH",
}
RDC_LABEL = "Remote Deposit Checks"
LB_LABEL = "LB 2999 Checks"
STATEMENT_LABEL = "ACH Statement Total"
ADJUSTED_LABEL = "Adjusted ACHs & Wires TOTAL"

SCALARS = {
    "umb_check_count": "H2",
    "remote_deposit_count": "I2",
    "remote_deposit_amount": RDC_LABEL,
    "lb2999_amount": LB_LABEL,
    "ach_statement_total": STATEMENT_LABEL,
    "ach_items": STATEMENT_LABEL + " items",
}


def plan_layout(plan):
    """
    Work out where every block sits for this particular day.

    At default sizes this reproduces the familiar layout exactly — loan pulls
    at 5-14, trade tickets at 17-36, TOTAL at 42 and so on. A block with more
    rows than usual grows, and everything below it moves down.
    """
    sizes = {k: max(SIZES[k], len(plan.get(k) or [])) for k in SIZES}
    L = {"sizes": sizes}

    L["loan_first"] = 5
    L["loan_last"] = L["loan_first"] + sizes["loan_pulls"] - 1
    L["loan_total"] = L["loan_last"] + 1

    L["ticket_first"] = L["loan_total"] + 2
    L["ticket_last"] = L["ticket_first"] + sizes["trade_tickets"] - 1
    L["ticket_total"] = L["ticket_last"] + 1

    L["ach_row"] = L["ticket_total"] + 2
    L["rdc_row"] = L["ach_row"] + 1
    L["lb_row"] = L["ach_row"] + 2
    L["total_row"] = L["lb_row"] + 1

    L["statement_row"] = L["total_row"] + 2
    L["pm_first"] = L["statement_row"] + 2
    L["pm_last"] = L["pm_first"] + sizes["pm_ach"] - 1
    L["wire_first"] = L["pm_last"] + 2
    L["wire_last"] = L["wire_first"] + sizes["wires"] - 1
    L["pulled_first"] = L["wire_last"] + 2
    L["pulled_last"] = L["pulled_first"] + sizes["pulled_ach"] - 1
    L["adjusted_row"] = L["pulled_last"] + 2

    L["blocks"] = {
        "loan_pulls": (L["loan_first"], L["loan_last"], "G"),
        "trade_tickets": (L["ticket_first"], L["ticket_last"], "F"),
        "pm_ach": (L["pm_first"], L["pm_last"], "C"),
        "wires": (L["wire_first"], L["wire_last"], "C"),
        "pulled_ach": (L["pulled_first"], L["pulled_last"], "C"),
    }
    return L


class Overflow(Exception):
    """Kept for callers that still catch it; blocks grow instead of failing."""


def check_capacity(plan):
    """Blocks grow to fit, so there is nothing left to reject."""
    return True


# --- styling --------------------------------------------------------------
ACCT = r'_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)'
CURRENCY = r'"$"#,##0.00_);\("$"#,##0.00\)'
YELLOW = PatternFill("solid", fgColor="FFFFFF00")   # keyed in by hand
GREEN = PatternFill("solid", fgColor="FF92D050")    # counted by formula
BLUE = PatternFill("solid", fgColor="FFDAEEF3")     # headline row shading


def _box(style):
    s = Side(style=style)
    return Border(left=s, right=s, top=s, bottom=s)


THIN_BOX = _box("thin")
MEDIUM_BOX = _box("medium")

WIDTHS = {"A": 4.57, "B": 44.0, "C": 32.14, "D": 24.29, "E": 24.29,
          "F": 20.0, "G": 17.43, "H": 8.57, "J": 8.57, "K": 8.57}

BLUE_CELLS = ["D1", "E1", "F1", "G1", "B2", "C2", "E2", "F2", "G2"]


_LINE_HEIGHT = 18.0
# Column widths are measured in characters of the workbook's default font
# (Calibri 11). The body cells are Arial 14, which is wider, so fewer
# characters fit on a line of column B than its width of 44 suggests.
_CHARS_PER_LINE = int(WIDTHS["B"] * 11 / 14)


def _wrapped_lines(text, chars=_CHARS_PER_LINE):
    """
    How many lines Excel will break `text` into inside column B.

    Excel wraps at spaces, splitting mid-word only when a single word is wider
    than the column. Counting the lines is what lets the row be made tall
    enough to show them: a row left at one line's height doesn't grow, it just
    hides everything after the first line.
    """
    lines, used = 1, 0
    for word in str(text or "").split():
        need = len(word) if used == 0 else len(word) + 1
        if used and used + need > chars:
            lines += 1
            used = len(word)
        else:
            used += need
        while used > chars:          # a single word wider than the column
            lines += 1
            used -= chars
    return max(1, lines)


def _put(ws, ref, value, *, size=14, bold=False, fill=None, fmt=None, align=None):
    cell = ws[ref]
    cell.value = value
    cell.font = Font(name="Arial", size=size, bold=bold)
    if fill:
        cell.fill = fill
    if fmt:
        cell.number_format = fmt
    if align:
        cell.alignment = Alignment(horizontal=align)
    return cell


def _finish(ws, L):
    """Apply layout and borders once, after every value is in place."""
    for col, width in WIDTHS.items():
        ws.column_dimensions[col].width = width

    # Trade ticket rows are deep enough for a wrapped two-line name; the ACH
    # blocks are tight.
    heights = {1: 55.5, 2: 26.25, 3: 20.25, 4: 15.75,
               L["loan_total"]: 18.0, L["loan_total"] + 1: 18.0,
               L["ticket_total"]: 18.0, L["ticket_total"] + 1: 18.0,
               L["ach_row"]: 18.0, L["rdc_row"]: 18.0, L["lb_row"]: 18.75,
               L["total_row"]: 42.75, L["total_row"] + 1: 21.75,
               L["statement_row"]: 17.25, L["adjusted_row"]: 15.75}
    # Loan pull and trade ticket labels carry the client name and wrap inside
    # column B, so each row is made as tall as its own text needs. Trade
    # tickets keep a two-line floor, which is what they have always been, so a
    # short name looks no different from before.
    for first, last, floor in (
            (L["loan_first"], L["loan_last"], _LINE_HEIGHT),
            (L["ticket_first"], L["ticket_last"], 2 * _LINE_HEIGHT)):
        for r in range(first, last + 1):
            lines = _wrapped_lines(ws[f"B{r}"].value)
            heights[r] = max(floor, _LINE_HEIGHT * lines)
    for first, last in ((L["pm_first"], L["pm_last"]),
                        (L["wire_first"], L["wire_last"]),
                        (L["pulled_first"], L["pulled_last"])):
        heights.update({r: 15.0 for r in range(first, last + 1)})
    for row, height in heights.items():
        ws.row_dimensions[row].height = height

    medium = ["D1:G1", "B2:G2", "B4:G4", f"B{L['total_row']}:G{L['total_row']}"]
    thin = ["H1:M1", "H2:M2", "B3:B3",
            f"B{L['loan_total']}:D{L['loan_total']}",
            f"B{L['ticket_total']}:D{L['ticket_total']}",
            f"B{L['ach_row']}:E{L['lb_row']}",
            f"A{L['loan_first']}:G{L['loan_last']}",
            f"A{L['ticket_first']}:G{L['ticket_last']}"]

    for ranges, border in ((medium, MEDIUM_BOX), (thin, THIN_BOX)):
        for span in ranges:
            for row in ws[span]:
                for cell in row:
                    cell.border = border

    for ref in BLUE_CELLS:
        ws[ref].fill = BLUE

    # One outline around the whole ACH reconciliation block, not a grid.
    thin_side, none = Side(style="thin"), Side()
    lower_first, lower_last = L["statement_row"], L["adjusted_row"]
    for row in range(lower_first, lower_last + 1):
        top = thin_side if row == lower_first else none
        bottom = thin_side if row == lower_last else none
        ws[f"B{row}"].border = Border(left=thin_side, right=none, top=top, bottom=bottom)
        ws[f"C{row}"].border = Border(left=none, right=thin_side, top=none, bottom=bottom)
    ws[f"E{lower_first}"].border = Border(right=thin_side)

    for cell in ws[1]:
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
    for row in ws[f"B1:B{lower_last}"]:
        for cell in row:
            keep = cell.alignment
            cell.alignment = Alignment(
                horizontal=keep.horizontal or ("center" if cell.row == 1 else None),
                vertical=keep.vertical or ("center" if cell.row == 1 else None),
                wrap_text=True)


def generate(plan, path):
    """Write a complete workbook for `plan` to `path`."""
    L = plan_layout(plan)
    deposit = plan["date"]
    if isinstance(deposit, str):
        deposit = _date.fromisoformat(deposit)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CONTROL SHEET"

    total_row = L["total_row"]
    statement = L["statement_row"]
    adjusted = L["adjusted_row"]

    pm = f"C{L['pm_first']}:C{L['pm_last']}"
    wire = f"C{L['wire_first']}:C{L['wire_last']}"
    pulled = f"C{L['pulled_first']}:C{L['pulled_last']}"
    tickets = f"F{L['ticket_first']}:F{L['ticket_last']}"

    # --- headline row ---
    for ref, text in (("D1", "DEPOSIT DATE"), ("E1", "PROCESS DATE"),
                      ("G1", "TOTAL # OF ITEMS")):
        _put(ws, ref, text, size=12, bold=True, align="center")
    for ref, text in (("H1", "UMB Checks"), ("I1", "Remote Deposit Checks"),
                      ("J1", "ACHs & Wires"), ("K1", "Trade Tickets"),
                      ("M1", "Total Debits")):
        cell = _put(ws, ref, text, size=11, bold=True, align="center")
        cell.font = Font(name="Calibri", size=11, bold=True)

    _put(ws, "B2", "TOTAL DEPOSIT", size=12, bold=True, fmt=CURRENCY, align="center")
    _put(ws, "C2", f"=D{total_row}", size=18, bold=True, fmt=ACCT, align="right")
    _put(ws, "D2", deposit, size=20, fill=YELLOW, fmt="mm/dd/yy", align="center")
    _put(ws, "E2", "=WORKDAY(D2,1)", size=18, fmt="mm/dd/yy", align="center")
    _put(ws, "G2", "=I2+J2+K2+H2", size=18, fmt="0", align="center")
    _put(ws, "H2", plan.get("umb_check_count"), fill=YELLOW, align="center")
    _put(ws, "I2", plan.get("remote_deposit_count"), fill=YELLOW, fmt="0", align="center")
    # Yellow is keyed by hand, green is counted by formula — the live sheet
    # uses the two fills to say which is which.
    _put(ws, "J2", f"=E{statement}-COUNTA({pm})+COUNTA({wire},{pulled})",
         fill=GREEN, align="center")
    _put(ws, "K2", f"=COUNTA({tickets})", fill=GREEN, align="center")
    _put(ws, "M2", f"=COUNTA({tickets},{pulled})", size=10, align="center")

    _put(ws, "B3", "CURRENT DATE LOCKBOX", size=11, bold=True)
    for ref, text in (("B4", "DESCRIPTION"), ("C4", "DEPOSIT DATE"),
                      ("D4", "TOTAL DEPOSIT"), ("E4", None), ("F4", None),
                      ("G4", None)):
        _put(ws, ref, text, size=11, bold=True, align="center")

    # --- loan pulls and trade tickets ---
    def entry_block(key, entries):
        first, last, amount_col = L["blocks"][key]
        for i, row in enumerate(range(first, last + 1)):
            item = entries[i] if i < len(entries) else None
            name = (item or {}).get("name", "") or ""
            _put(ws, f"A{row}", i + 1)
            _put(ws, f"B{row}", f"{LABELS[key]} ({name})", fmt="00000000", align="left")
            _put(ws, f"C{row}", "=$D$2", fmt="mm/dd/yy", align="left")
            _put(ws, f"D{row}", f"={amount_col}{row}", fmt=ACCT, align="left")
            for col in ("E", "F", "G"):
                if col != amount_col:
                    _put(ws, f"{col}{row}", None, fmt=ACCT, align="left")
            _put(ws, f"{amount_col}{row}", (item or {}).get("amount"),
                 fill=YELLOW, fmt=ACCT, align="left")

    entry_block("loan_pulls", plan.get("loan_pulls") or [])
    _put(ws, f"B{L['loan_total']}", "TOTAL LOAN PULLS", bold=True,
         fmt="00000000", align="right")
    _put(ws, f"C{L['loan_total']}", None, fmt="mm-dd-yy", align="left")
    _put(ws, f"D{L['loan_total']}", f"=SUM(D{L['loan_first']}:D{L['loan_last']})",
         bold=True, fmt=ACCT, align="left")

    entry_block("trade_tickets", plan.get("trade_tickets") or [])
    _put(ws, f"B{L['ticket_total']}", "TOTAL TRADE TICKETS", bold=True,
         fmt="00000000", align="right")
    _put(ws, f"C{L['ticket_total']}", None, fmt="mm-dd-yy", align="left")
    _put(ws, f"D{L['ticket_total']}",
         f"=SUM(D{L['ticket_first']}:D{L['ticket_last']})",
         bold=True, fmt=ACCT, align="left")

    # --- the three single-line entries ---
    for row, label, value, keyed in (
            (L["ach_row"], "ACH & Wires pushed & pulled", f"=C{adjusted}", False),
            (L["rdc_row"], RDC_LABEL, plan.get("remote_deposit_amount"), True),
            (L["lb_row"], LB_LABEL, plan.get("lb2999_amount"), True)):
        _put(ws, f"B{row}", label, fmt="00000000", align="left")
        _put(ws, f"C{row}", "=$D$2", fmt="mm/dd/yy", align="left")
        _put(ws, f"D{row}", f"=E{row}", bold=True, fmt=ACCT, align="left")
        _put(ws, f"E{row}", value, fill=YELLOW if keyed else None,
             fmt=ACCT, align="right")

    _put(ws, f"B{total_row}", "TOTAL", size=11, bold=True, align="left")
    _put(ws, f"C{total_row}", None, size=11, fmt="mm-dd-yy", align="left")
    for col, formula in (
            ("D", f"=SUM(D{L['loan_first']}:D{L['loan_last']},"
                  f"D{L['ticket_first']}:D{L['ticket_last']},"
                  f"D{L['ach_row']}:D{L['lb_row']})"),
            ("E", f"=SUM(E4:E{L['lb_row']})"),
            ("F", f"=SUM(F{L['loan_first']}:F{L['lb_row']})"),
            ("G", f"=SUM(G{L['loan_first']}:G{L['lb_row']})")):
        _put(ws, f"{col}{total_row}", formula, bold=True, fmt=ACCT, align="left")

    # --- ACH reconciliation ---
    _put(ws, f"B{statement}", STATEMENT_LABEL, size=11, bold=True)
    _put(ws, f"C{statement}", plan.get("ach_statement_total"), size=12,
         fill=YELLOW, fmt=ACCT)
    _put(ws, f"D{statement}", "Total ACH UMB items:", size=10, align="right")
    _put(ws, f"E{statement}", plan.get("ach_items"), size=12, fill=YELLOW)

    for key in ("pm_ach", "wires", "pulled_ach"):
        first, last, _col = L["blocks"][key]
        amounts = plan.get(key) or []
        for i, row in enumerate(range(first, last + 1)):
            _put(ws, f"A{row}", i + 1, size=10)
            _put(ws, f"B{row}", LABELS[key], size=10)
            _put(ws, f"C{row}", amounts[i] if i < len(amounts) else None,
                 size=12, fill=YELLOW, fmt=ACCT)

    _put(ws, f"B{adjusted}", ADJUSTED_LABEL, size=12, bold=True)
    _put(ws, f"C{adjusted}", f"=C{statement}-SUM({pm})+SUM({wire})+SUM({pulled})",
         size=12, bold=True, fmt=ACCT)

    _finish(ws, L)

    # The sheet's own cross-foot checks. These cells carry no fill of their
    # own — a conditional format paints them green only while the total ties.
    # Differential styles colour from bgColor, not fgColor: setting fgColor
    # writes a rule that fires but shows nothing.
    good_fill = PatternFill(bgColor="FFC6EFCE", fill_type="solid")
    good_font = Font(color="FF006100")
    for col, expected in (
            ("D", f"SUM($D${L['loan_total']},$D${L['ticket_total']},"
                  f"$D${L['ach_row']}:$D${L['lb_row']})"),
            ("E", f"SUM($D${L['ach_row']}:$D${L['lb_row']})"),
            ("F", f"$D${L['ticket_total']}"),
            ("G", f"$D${L['loan_total']}")):
        ws.conditional_formatting.add(f"{col}{total_row}", CellIsRule(
            operator="equal", formula=[expected], fill=good_fill, font=good_font))

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
    m["B2"].number_format = ACCT
    m["G2"].number_format = "0"

    wb.save(path)
    return path


# --- reading and comparing ------------------------------------------------
def read_values(path):
    """
    Pull the input values back out of a workbook.

    Rows are found by their labels in column B rather than by fixed numbers,
    so a sheet whose blocks were sized differently — ours on a busy day, or
    one someone extended by hand — still reads correctly.
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["CONTROL SHEET"]

    labels = {}
    for row in range(1, ws.max_row + 1):
        text = ws[f"B{row}"].value
        if isinstance(text, str) and text.strip():
            # "CURRENT DATE TRADE TICKET (Doe, Jane)" -> the label alone.
            stem = re.sub(r"\s*\([^()]*\)\s*$", "", text.strip())
            labels.setdefault(stem, []).append(row)

    def block(key):
        col = AMOUNT_COL[key]
        return [ws[f"{col}{r}"].value for r in labels.get(LABELS[key], [])
                if ws[f"{col}{r}"].value not in (None, "")]

    def single(label, col):
        rows = labels.get(label, [])
        return ws[f"{col}{rows[0]}"].value if rows else None

    out = {
        "date": ws["D2"].value,
        "umb_check_count": ws["H2"].value,
        "remote_deposit_count": ws["I2"].value,
        "remote_deposit_amount": single(RDC_LABEL, "E"),
        "lb2999_amount": single(LB_LABEL, "E"),
        "ach_statement_total": single(STATEMENT_LABEL, "C"),
        "ach_items": single(STATEMENT_LABEL, "E"),
        "total_deposit": single("TOTAL", "D"),
        "adjusted_ach": single(ADJUSTED_LABEL, "C"),
    }
    for key in SIZES:
        out[key] = block(key)
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

    Blocks are compared as the values actually present, not cell by cell, so
    a stray blank row — or a block sized differently from ours — isn't
    reported as a difference. Only the numbers matter.
    """
    theirs = read_values(existing_path)
    rows = []

    for key, where in SCALARS.items():
        ours = plan.get(key)
        rows.append({
            "what": key, "cell": where,
            "ours": ours, "theirs": theirs.get(key),
            "same": _same(ours, theirs.get(key)),
        })

    for key in SIZES:
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
            "what": key, "cell": LABELS[key],
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
