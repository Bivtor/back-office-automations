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

# Column boundaries in PDF points (x0), as measured on a 612pt-wide portrait
# page. They are scaled to whatever width the file actually has — see
# _columns_for. The report has been exported both portrait (612pt) and
# landscape (792pt), and the columns sit at the same fractions of the page
# either way, so the ratios hold and the absolute points do not.
COLUMNS = {
    "date":       (55, 125),   # Payment Date (line 1) / Send Date (line 2)
    "payment_no": (125, 200),  # Payment No. (line 1) / Name-Reference (line 2+)
    "status":     (200, 270),  # Status (line 1) / Confirmation No. (line 2+)
    "account":    (270, 350),  # Co. Account / Co. Account Identifier
    "type":       (350, 430),  # Type / Created By Template
    "recipient":  (430, 500),  # Recipient
    "amount":     (500, 578),  # Recipient Amount (Items)
}

REFERENCE_WIDTH = 612.0   # the width those bounds were measured on

DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{4}$")
AMOUNT_RE = re.compile(r"^[\d,]+\.\d{2}$")
ITEMS_RE = re.compile(r"^\((\d+)\)$")
# "LB 419300", "LB416525" and a bare "LB" all turn up, so the number is
# optional and the space is too. Case is ignored for the same reason the loan
# matcher uses a prefix: this text is the bank's to change, not ours.
LB_NUMBER_RE = re.compile(r"^LB([\s\-]*\d+)?$", re.IGNORECASE)


# Statuses that still land on the control sheet. A Scheduled payment is one the
# bank has accepted and dated but not yet settled, so it belongs on the sheet
# the same as a Completed one. Anything else — Deleted above all — is a payment
# that isn't happening, and must stay out.
COUNTED_STATUSES = {"completed", "scheduled"}


def _is_loan(ref):
    return ref.strip().upper().startswith("LOAN")


def _is_lockbox_ach(ref):
    ref = ref.strip()
    return ref.upper() == "LOCKBOX" or bool(LB_NUMBER_RE.match(ref))

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


def _columns_for(width):
    """
    The column bounds for a page of this width.

    Chrome lays the report out to whatever paper it is given, so a landscape
    export puts every column at the same fraction of the page but a third
    further across in points. Scaling by the width reads both. Without it a
    792pt page leaves every amount past the right edge of the amount column,
    so every row parses with no amount and would count as zero.
    """
    scale = (width or REFERENCE_WIDTH) / REFERENCE_WIDTH
    return {name: (lo * scale, hi * scale) for name, (lo, hi) in COLUMNS.items()}


def _column_of(x0, columns):
    for name, (lo, hi) in columns.items():
        if lo <= x0 < hi:
            return name
    return None


def _parse_page(page):
    columns = _columns_for(page.width)
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
        if columns["amount"][0] <= w["x0"] < columns["amount"][1]
        and AMOUNT_RE.match(w["text"])
    )
    # A record begins on any line carrying a date in the leftmost data column.
    anchors = sorted(
        w["top"] for w in words
        if columns["date"][0] <= w["x0"] < columns["date"][1]
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
            col = _column_of(w["x0"], columns)
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


def categorize(rows, source, added=None):
    """
    Split rows into the control-sheet buckets for `source`
    ("loans_tt" or "ach").

    `added` is {payment_no: bucket} for rows the user dragged into a bucket by
    hand. It overrides both tests below — status and Name/Reference — because
    the whole point of the override is the day the rules read a row wrongly.
    Buckets that don't belong to this source are ignored rather than trusted,
    so a stale key can't invent a category.

    Returns dict with:
      buckets   - {category: [rows]}   rows we will use
      skipped   - [rows]               status isn't one we count
      unmatched - [rows]               countable, but Name/Reference not ours
    """
    rules = CATEGORIES[source]
    forced = {no: b for no, b in (added or {}).items() if b in rules}
    buckets = {name: [] for name in rules}
    skipped, unmatched = [], []

    for row in rows:
        chosen = forced.get((row.get("payment_no") or "").strip())
        if chosen:
            buckets[chosen].append(row)
            continue
        if row["status"].strip().lower() not in COUNTED_STATUSES:
            skipped.append(row)
            continue
        for name, matches in rules.items():
            if matches(row["name_ref"]):
                buckets[name].append(row)
                break
        else:
            unmatched.append(row)

    return {"buckets": buckets, "skipped": skipped, "unmatched": unmatched}


def row_key(bucket, index, row):
    """
    A stable handle for one row, so the browser can say "leave this one out".

    Payment numbers are unique per row and survive a re-pull of the same PDF,
    which position does not — the afternoon file has more rows than the morning
    one, so anything index-based would silently shift onto its neighbour. The
    positional form is only a fallback for a row whose number didn't read.
    """
    no = (row.get("payment_no") or "").strip()
    return no or f"{bucket}#{index}"


def keep(bucket, rows, dropped):
    """Rows the user hasn't struck out."""
    dropped = set(dropped or ())
    return [r for i, r in enumerate(rows) if row_key(bucket, i, r) not in dropped]


def totals(rows):
    return {
        "count": len(rows),
        "items": sum(r.get("items") or 1 for r in rows),
        "sum": round(sum(r["amount"] or 0 for r in rows), 2),
    }
