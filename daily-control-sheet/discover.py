"""
Locate the four daily source files.

Paths shift a little day to day, so nothing here assumes an exact filename.
Each source globs its folder, scores every file it finds, and returns the best
guess plus every candidate so the UI can offer an override.
"""
import os
import re
from datetime import date

# Change this if the share is mapped to a different letter.
BASE = os.environ.get("CONTROL_SHEET_BASE", r"Q:\CORE\PURCHASE TRADING")

MONTHS = ["January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December"]


def date_tokens(d):
    """Every plausible way this date appears in a filename. Always includes a
    year — without one, '7-31' matches 7-31-23 as readily as 7-31-26."""
    m, dd, yy, yyyy = d.month, d.day, d.strftime("%y"), d.year
    numbers = {
        (str(m), str(dd)),
        (f"{m:02d}", str(dd)),
        (str(m), f"{dd:02d}"),
        (f"{m:02d}", f"{dd:02d}"),
    }
    tokens = set()
    for a, b in numbers:
        for sep in ("-", ".", "_", " "):
            for year in (yy, str(yyyy)):
                tokens.add(f"{a}{sep}{b}{sep}{year}")
    return tokens


def umb_folder():
    return os.path.join(BASE, "UMB Confirms", "UMB Reports Daily Totals")


def lockbox_folder(d):
    return os.path.join(
        BASE, "Lockbox Deposit Images",
        f"{d.year} Lockbox Deposit Images",
        f"{d.month:02d} {MONTHS[d.month - 1]} Lockbox Deposit Images",
        f"{d.day:02d}",
    )


def _score(filename, d, must_have, must_not_have=()):
    name = filename.lower()
    if any(bad in name for bad in must_not_have):
        return 0
    if not all(any(v in name for v in variants) for variants in must_have):
        return 0
    score = 10
    if any(tok in name for tok in date_tokens(d)):
        score += 10
    return score


def _find(folder, d, must_have, must_not_have=(), extensions=(".pdf",),
          date_in_folder=False):
    """
    date_in_folder=True for sources whose folder path already pins the date
    (the Lockbox tree nests year/month/day), so their filenames never carry it.
    """
    result = {"folder": folder, "path": None, "candidates": [],
              "confident": False, "error": None}
    if not os.path.isdir(folder):
        result["error"] = f"Folder not found: {folder}"
        return result

    scored = []
    for entry in sorted(os.listdir(folder)):
        full = os.path.join(folder, entry)
        if not os.path.isfile(full):
            continue
        if extensions and not entry.lower().endswith(extensions):
            continue
        scored.append((_score(entry, d, must_have, must_not_have), entry, full))

    result["candidates"] = [
        {"name": n, "path": p, "score": s}
        for s, n, p in sorted(scored, key=lambda x: (-x[0], x[1]))
    ]
    best = [c for c in result["candidates"] if c["score"] > 0]
    if best:
        result["path"] = best[0]["path"]
        # Score 20 means the filename carried today's date. Score 10 means only
        # the keywords matched — fine when the folder already pins the date,
        # otherwise it could easily be another day's file.
        result["confident"] = best[0]["score"] >= 20 or date_in_folder
        if not result["confident"]:
            result["error"] = ("Filename doesn't contain this date — check it's "
                               "the right day before reading.")
    else:
        result["error"] = "No file in this folder matched. Pick one below."
    return result


def parse_check_filename(name):
    """
    Read the check count and dollar total straight off the filename. Seen so far:

        58 Check totaling $1,446,953.10.xls
        32 Check totaling $2,647,088.xls      (no cents)
        39 Checks 2338512.27.xls              (no "totaling", no $ or commas)

    The count leads, the amount trails. Everything between them varies, so
    neither is matched on the wording around it. The file is never opened.
    """
    stem = os.path.splitext(name)[0]
    count = re.match(r"\s*(\d+)\b", stem)
    # Anchored at the end so the leading count can't be mistaken for the total.
    amount = re.search(r"\$?\s*([\d,]+(?:\.\d{2})?)\s*(?:\(\d+\))?\s*$", stem)
    value = _amount_or_none(amount.group(1)) if amount else None
    # A bare count with no amount ("39 Checks") would otherwise read its own
    # count back as the total.
    if count and amount and amount.start(1) < count.end(1):
        value = None
    return {
        "count": int(count.group(1)) if count else None,
        "amount": value,
    }


def _amount_or_none(text):
    try:
        return float(str(text).replace(",", ""))
    except (TypeError, ValueError):
        return None


def parse_leading_count(name):
    m = re.match(r"\s*(\d+)\b", name)
    return int(m.group(1)) if m else None


def output_path(d):
    """
    Where today's workbook belongs:
      ...\\UMB Lockbox Deposits-Daily Spreadsheets\\CONTROL SHEET 2026\\08- August\\4.xlsx

    Note this tree names months differently from the Lockbox tree ("08- August"
    rather than "08 August Lockbox Deposit Images") and the filename has no
    leading zero. Both the year and month folders are matched loosely so a
    stray space or a missing hyphen doesn't break the run.
    """
    root = os.path.join(BASE, "UMB Lockbox Deposits-Daily Spreadsheets")
    month_name = MONTHS[d.month - 1].lower()
    result = {"folder": None, "path": None, "exists": False, "error": None}

    year_dir = None
    if os.path.isdir(root):
        for entry in os.listdir(root):
            if str(d.year) in entry and os.path.isdir(os.path.join(root, entry)):
                year_dir = os.path.join(root, entry)
                break
    if not year_dir:
        result["error"] = f"No folder for {d.year} under {root}"
        return result

    month_dir = None
    for entry in sorted(os.listdir(year_dir)):
        low = entry.lower()
        if os.path.isdir(os.path.join(year_dir, entry)) and (
                month_name in low or low.startswith(f"{d.month:02d}")):
            month_dir = os.path.join(year_dir, entry)
            break
    if not month_dir:
        result["error"] = f"No folder for {MONTHS[d.month - 1]} under {year_dir}"
        return result

    result["folder"] = month_dir
    # Prefer a file that's already there, whichever way the day is written.
    for stem in (str(d.day), f"{d.day:02d}"):
        candidate = os.path.join(month_dir, f"{stem}.xlsx")
        if os.path.isfile(candidate):
            result["path"] = candidate
            result["exists"] = True
            return result
    result["path"] = os.path.join(month_dir, f"{d.day}.xlsx")
    return result


def locate_all(d=None):
    d = d or date.today()
    umb = umb_folder()
    lock = lockbox_folder(d)

    sources = {
        "loans_tt": _find(
            umb, d,
            must_have=[("loan",), ("tt", "trade")],
            extensions=(".pdf",),
        ),
        "ach": _find(
            umb, d,
            must_have=[("ach",)],
            must_not_have=("loan",),
            extensions=(".pdf",),
        ),
        "afternoon_report": _find(
            lock, d,
            must_have=[("ach",), ("afternoon",)],
            extensions=(".xls", ".xlsx"),
            date_in_folder=True,
        ),
        "check_total": _find(
            lock, d,
            # "Check totaling $X" and "Checks X" both turn up, so match on
            # "check" alone and keep the ACH report out of the running.
            must_have=[("check",)],
            must_not_have=("ach", "afternoon"),
            extensions=(".xls", ".xlsx"),
            date_in_folder=True,
        ),
    }

    # The check file's name carries the UMB check count and dollar total.
    ct = sources["check_total"]
    ct["parsed"] = parse_check_filename(os.path.basename(ct["path"])) if ct["path"] else {}

    ar = sources["afternoon_report"]
    ar["parsed"] = (
        {"leading_count": parse_leading_count(os.path.basename(ar["path"]))}
        if ar["path"] else {}
    )

    return {"date": d.isoformat(), "base": BASE, "sources": sources,
            "output": output_path(d)}
