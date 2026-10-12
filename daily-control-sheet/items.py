	"""
Parse a Chrome/Skia-exported "Payment Center" PDF into categorized rows.

No OCR, no AI. Column positions are read from the PDF's own coordinates.
Requires the PDF to have been saved with Chrome's "Save as PDF" destination,
NOT "Microsoft Print to PDF" (which outlines the glyphs and destroys the
text layer).
"""
import re
from collections import defaultdict

import pdfplumber

# Column boundaries in PDF points (x0). Letter page is 612pt wide.
# Verified stable across the 2024 and 2026 exports of this report.
COLUMNS = {
    "date":       (55, 125),   # Payment Date (line 1) / Send Date (line 2)
    "payment_no": (125, 200),  # Payment No. (line 1) / Name-Reference (line 2+)
    "status":     (200, 270),  # Status (line 1) / Confirmation No. (line 2+)
    "account":    (270, 350),  # Co. Account / Co. Account Identifier
    "type":       (350, 430),  # Type / Created By Template
    "recipient":  (430, 500),  # Recipient
    "amount":     (500, 578),  # Recipient Amount (Items)
}

DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{4}$")
AMOUNT_RE = re.compile(r"^[\d,]+\.\d{2}$")
ITEMS_RE = re.compile(r"^\((\d+)\)$")
LB_NUMBER_RE = re.compile(r"^LB \d+$")

# The loan reference is written "Loan" some days and "Loan Pull" others, so
# match on the prefix rather than an exact string. Nothing else in this report
# starts with "loan", and anything unexpected still lands in `unmatched` where
# it gets reported rather than silently dropped.
def _is_loan(ref):
    return ref.strip().upper().startswith("LOAN")


def _is_lockbox_ach(ref):
    ref = ref.strip()
    return ref == "Lockbox" or bool(LB_NUMBER_RE.match(ref))

# Chrome subsets the embedded font; the "fi" ligature arrives as U+0000
# and icomoon glyphs as U+E900.
LIGATURES = {"\x00": "fi", "\ue900": ""}

# Which Name/Reference values belong to which control-sheet block.
CATEGORIES = {
    "loans_tt": {
        "loan_pulls":    _is_loan,
        "trade_tickets": lambda ref: ref.strip() == "Lockbox TT",
    },
    "ach": {
        "ach_lockbox":   _is_lockbox_ach,
    },
}


def _clean(text):
    for bad, good in LIGATURES.items():
        text = text.replace(bad, good)
    return text.strip()


def _column_of(x0):
    for name, (lo, hi) in COLUMNS.items():
        if lo <= x0 < hi:
            return name
    return None


def _parse_page(page):
    words = []
    for w in page.extract_words():
        w["text"] = _clean(w["text"])
        if w["text"]:
            words.append(w)

    # Every record has exactly one amount. Counting the amount-shaped tokens
    # independently of the records gives us a way to notice rows that got
    # absorbed into their neighbour — see scan().
    amount_tokens = sum(
        1 for w in words
        if COLUMNS["amount"][0] <= w["x0"] < COLUMNS["amount"][1]
        and AMOUNT_RE.match(w["text"])
    )
    # A record begins on any line carrying a date in the leftmost data column.
    anchors = sorted(
        w["top"] for w in words
        if COLUMNS["date"][0] <= w["x0"] < COLUMNS["date"][1]
        and DATE_RE.match(w["text"])
    )
    # Two dates per record (payment date, send date) sit ~8pt apart.
    starts = []
    for t in anchors:
        if not starts or t - starts[-1] > 20:
            starts.append(t)
    if not starts:
        return [], amount_tokens

    bounds = starts + [page.height]
    records = []

    for i, top in enumerate(starts):
        band = [w for w in words if top - 2 <= w["top"] < bounds[i + 1] - 2]

        cells = defaultdict(list)
        for w in sorted(band, key=lambda w: (round(w["top"]), w["x0"])):
            col = _column_of(w["x0"])
            if col:
                cells[col].append(w)

        on_line1 = {id(w) for w in band if abs(w["top"] - top) < 4}

        def join(col, first_only=False, skip_first=False):
            out = [
                w["text"] for w in cells[col]
                if not (first_only and id(w) not in on_line1)
                and not (skip_first and id(w) in on_line1)
            ]
            # Re-join names hyphen-wrapped across lines ("Lopez- Roman").
            return re.sub(r"(?<=\w)- (?=\w)", "-", " ".join(out))

        dates = [w["text"] for w in cells["date"]]
        amounts = [w["text"] for w in cells["amount"] if AMOUNT_RE.match(w["text"])]
        counts = [ITEMS_RE.match(w["text"]) for w in cells["amount"]]
        counts = [int(m.group(1)) for m in counts if m]

        records.append({
            "payment_date": dates[0] if dates else "",
            "send_date":    dates[1] if len(dates) > 1 else "",
            "payment_no":   join("payment_no", first_only=True),
            "name_ref":     join("payment_no", skip_first=True),
            "status":       join("status", first_only=True),
            "recipient":    join("recipient"),
            "amount":       float(amounts[0].replace(",", "")) if amounts else None,
            # A batch row carries many items: "1,305,586.94 USD (3844)".
            "items":        counts[0] if counts else 1,
        })

    return records, amount_tokens


def scan(path):
    """
    Read the PDF and check it for rows we might have lost.

    Records are found by anchoring on the payment date in the leftmost column.
    A row with no readable date would be absorbed into the record above it, and
    two rows closer together than the record spacing could merge — both would
    vanish without a trace. Counting amount-shaped tokens separately catches
    either case, because every record carries exactly one amount.
    """
    with pdfplumber.open(path) as pdf:
        if not pdf.pages[0].chars:
            raise ValueError(
                "No text layer in this PDF. It was saved with 'Microsoft Print "
                "to PDF'. Re-save it using Chrome's 'Save as PDF' destination."
            )
        rows, expected = [], 0
        for page in pdf.pages:
            page_rows, tokens = _parse_page(page)
            rows.extend(page_rows)
            expected += tokens

    no_amount = [r["payment_no"] or "(no payment no.)"
                 for r in rows if r["amount"] is None]
    problems = []
    if len(rows) != expected:
        problems.append(
            f"Found {len(rows)} rows but {expected} amounts — "
            f"{abs(expected - len(rows))} row(s) may have been merged or missed.")
    if no_amount:
        problems.append(
            f"{len(no_amount)} row(s) have no readable amount and would count "
            f"as zero: {', '.join(no_amount[:5])}")

    return {"rows": rows, "row_count": len(rows), "amounts_found": expected,
            "problems": problems, "ok": not problems}


def read_pdf(path):
    """Every row in the PDF, uncategorized and unfiltered."""
    return scan(path)["rows"]


def categorize(rows, source):
    """
    Split rows into the control-sheet buckets for `source`
    ("loans_tt" or "ach").

    Returns dict with:
      buckets   - {category: [rows]}   rows we will use
      skipped   - [rows]               status is not Completed
      unmatched - [rows]               Completed, but Name/Reference not ours
    """
    rules = CATEGORIES[source]
    buckets = {name: [] for name in rules}
    skipped, unmatched = [], []

    for row in rows:
        if row["status"].strip().lower() != "completed":
            skipped.append(row)
            continue
        for name, matches in rules.items():
            if matches(row["name_ref"]):
                buckets[name].append(row)
                break
        else:
            unmatched.append(row)

    return {"buckets": buckets, "skipped": skipped, "unmatched": unmatched}


def totals(rows):
    return {
        "count": len(rows),
        "items": sum(r.get("items") or 1 for r in rows),
        "sum": round(sum(r["amount"] or 0 for r in rows), 2),
    }
