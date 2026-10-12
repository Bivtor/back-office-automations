"""
Parse the UMB "ACHs Afternoon Report" (legacy .xls).

The report is heavily merged, but xlrd returns each merged block's value in its
top-left cell, so plain (row, col) reads work and merges can be ignored.

Layout, by column:
  B  description lines and the free-text detail under each entry
  C  section total labels ("ACH Credit Received Total:")
  H  item count on total rows ("  39 Item(s)")
  I  0-Day amount
  O  total amount

An ACH entry is a row with "ACH Credit Received" in B; its amount sits in I on
that same row and its name on the row BELOW.

A wire is a block starting with "Incoming Money Transfer" in B, running until
the next wire or a section total. The sender is on the block's "ORG=" line —
except some wires have no ORG= line at all, so we fall back to OBI= then BNF=.
"""
import re

import xlrd

# Wires from these senders are excluded by default. The user can override any
# of them in the UI, so this list is a starting position, not a rule.
EXCLUDE_BY_DEFAULT = ["NSCC", "CHARLES SCHWAB", "SCHWAB", "BANK OF THE SIERRA"]

HOUSE_PREFIX = "INTERNAL_CO"

ACH_ROW = "ACH Credit Received"
ACH_TOTAL = "ACH Credit Received Total:"
WIRE_ROW = "Incoming Money Transfer"
WIRE_TOTAL = "Incoming Money Transfer Total:"
LOCKBOX_TOTAL = "Lockbox Deposit Total:"

B, C, H, I, O = 1, 2, 7, 8, 14


def _money(text):
    text = str(text or "").strip().replace(",", "")
    try:
        return float(text)
    except ValueError:
        return None


def _items(text):
    m = re.search(r"(\d+)\s*Item", str(text or ""))
    return int(m.group(1)) if m else None


def _sender(lines):
    """Pull the counterparty out of a wire's detail lines."""
    for prefix in ("ORG=", "OBI=", "BNF="):
        for line in lines:
            if line.upper().startswith(prefix):
                text = line[len(prefix):].strip()
                # "SOME VENDOR LLC*1 MAIN ST*ANYTOWN NY US*"
                parts = [p.strip() for p in text.split("*") if p.strip()]
                if parts:
                    return {"name": parts[0], "detail": " · ".join(parts),
                            "found_on": prefix.rstrip("=")}
    return {"name": "(no sender line)", "detail": "", "found_on": None}


def read(path):
    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_index(0)

    def cell(r, c):
        try:
            return str(sheet.cell_value(r, c)).strip()
        except IndexError:
            return ""

    ach_rows, house, wires = [], [], []
    stated = {}

    r = 0
    while r < sheet.nrows:
        b, c = cell(r, B), cell(r, C)

        if c == ACH_TOTAL:
            stated["ach_total"] = _money(cell(r, I))
            stated["ach_items"] = _items(cell(r, H))
        elif c == WIRE_TOTAL:
            stated["wire_total"] = _money(cell(r, I))
            stated["wire_items"] = _items(cell(r, H))
        elif c == LOCKBOX_TOTAL:
            stated["lockbox_total"] = _money(cell(r, O))
            stated["lockbox_items"] = _items(cell(r, H))

        if b == ACH_ROW:
            entry = {"row": r + 1, "name": cell(r + 1, B), "amount": _money(cell(r, I))}
            ach_rows.append(entry)
            if entry["name"].upper().startswith(HOUSE_PREFIX):
                house.append(entry)

        elif b == WIRE_ROW:
            amount = _money(cell(r, I))
            lines, k = [], r + 1
            while k < sheet.nrows:
                nb, nc = cell(k, B), cell(k, C)
                if nb == WIRE_ROW or nc.endswith("Total:"):
                    break
                if nb:
                    lines.append(nb)
                k += 1

            sender = _sender(lines)
            blob = " ".join(lines).upper()
            hits = [term for term in EXCLUDE_BY_DEFAULT if term in blob]
            wires.append({
                "row": r + 1,
                "amount": amount,
                "sender": sender["name"],
                "detail": sender["detail"],
                "found_on": sender["found_on"],
                # False means "leave it out" — the UI can flip this.
                "include": not hits,
                "excluded_for": hits[0] if hits else None,
            })
            r = k - 1

        r += 1

    return {
        "ach": {
            "total": stated.get("ach_total"),
            "items": stated.get("ach_items"),
            "row_count": len(ach_rows),
            "row_sum": round(sum(e["amount"] or 0 for e in ach_rows), 2),
        },
        "house": house,
        "house_sum": round(sum(e["amount"] or 0 for e in house), 2),
        "wires": wires,
        "stated": stated,
        "checks": {
            "ach_rows_match_total":
                stated.get("ach_items") == len(ach_rows)
                and abs(sum(e["amount"] or 0 for e in ach_rows)
                        - (stated.get("ach_total") or 0)) < 0.005,
            "wire_rows_match_total":
                stated.get("wire_items") == len(wires)
                and abs(sum(w["amount"] or 0 for w in wires)
                        - (stated.get("wire_total") or 0)) < 0.005,
        },
    }
