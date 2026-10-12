"""
Locate the four daily source files.

Paths shift a little day to day, so nothing here assumes an exact filename.
Each source globs its folder, scores every file it finds, and returns the best
guess plus every candidate so the UI can offer an override.
"""
import os
import re
from datetime import date

import afternoon

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


def downloads_folder():
    """
    Where the PDFs land when they're saved out of the browser first.

    The share copy doesn't appear until the afternoon, so on a 3pm run the only
    copy of a PDF may be the one in Downloads. Overridable for a machine that
    puts it somewhere else.
    """
    return os.environ.get(
        "CONTROL_SHEET_DOWNLOADS",
        os.path.join(os.path.expanduser("~"), "Downloads"))


def lockbox_folder(d):
    return os.path.join(
        BASE, "Lockbox Deposit Images",
        f"{d.year} Lockbox Deposit Images",
        f"{d.month:02d} {MONTHS[d.month - 1]} Lockbox Deposit Images",
        f"{d.day:02d}",
    )


# Folder listings for the lookup in progress. Every folder here gets read
# twice — the two PDFs share the UMB folder and share Downloads, the report and
# the check file share the day folder — and over a VPN a second read of the
# same folder is a second round trip for nothing.
_listings = {}


def _files_in(folder):
    """
    Names and paths of the files in a folder, in one pass.

    os.scandir rather than listdir + isfile: scandir hands back what the
    directory enumeration already told it, so asking whether each entry is a
    file costs nothing. listdir + os.path.isfile asks the server again for
    every single name, which is fine on a local disk and ruinous over a VPN —
    a folder holding a couple of years of daily reports turns one lookup into
    thousands of round trips, and the whole app sits there waiting.
    """
    if folder in _listings:
        return _listings[folder]
    out = []
    try:
        with os.scandir(folder) as it:
            for entry in it:
                try:
                    if entry.is_file():
                        out.append((entry.name, entry.path))
                except OSError:
                    continue
    except OSError:
        out = None                      # missing, or not reachable
    _listings[folder] = out
    return out


def _score(filename, d, must_have, must_not_have=(), prefer=()):
    name = filename.lower()
    if any(bad in name for bad in must_not_have):
        return 0
    if not all(any(v in name for v in variants) for variants in must_have):
        return 0
    score = 10
    if any(tok in name for tok in date_tokens(d)):
        score += 10
    # A tie-break, not a filter. Both ACH reports are valid input; this only
    # decides which is offered first when both are sitting in the folder.
    if any(p in name for p in prefer):
        score += 5
    return score


def _find(folder, d, must_have, must_not_have=(), extensions=(".pdf",),
          date_in_folder=False, prefer=()):
    """
    date_in_folder=True for sources whose folder path already pins the date
    (the Lockbox tree nests year/month/day), so their filenames never carry it.
    """
    result = {"folder": folder, "path": None, "candidates": [],
              "confident": False, "error": None}
    files = _files_in(folder)
    if files is None:
        result["error"] = f"Folder not found: {folder}"
        return result

    scored = []
    for entry, full in sorted(files):
        if extensions and not entry.lower().endswith(extensions):
            continue
        scored.append((_score(entry, d, must_have, must_not_have, prefer), entry, full))

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


def _find_in(folders, d, **kw):
    """
    Run _find over several folders and merge the results.

    Folders are given in priority order, the share first and Downloads as the
    fallback. Every candidate from every folder is offered, so a file sitting
    in Downloads is picked up when the share copy isn't there yet.

    Ranking is by filename score first, then by folder. Score has to lead,
    because the score is what says the file carries today's date: if the share
    is still holding yesterday's PDF and today's is in Downloads, taking the
    share copy on principle would silently build the sheet from the wrong day.
    Folder decides everything after that, so whenever a share file and a
    Downloads file are equally good the share one is the default. That tie is
    common and used to be settled by whichever filename sorted first, which is
    no basis for anything.

    Passing over a share file is worth saying out loud, so `note` carries it.
    A folder that doesn't exist is skipped; the first folder's error is kept
    only if nothing was found anywhere.
    """
    folders = [f for f in folders if f]
    results = [_find(f, d, **kw) for f in folders]

    candidates = []
    for rank, r in enumerate(results):
        for c in r["candidates"]:
            if c["score"] > 0:
                candidates.append({**c, "rank": rank, "folder": folders[rank],
                                   "confident": r["confident"]})
    candidates.sort(key=lambda c: (-c["score"], c["rank"], c["name"]))

    merged = dict(results[0])
    merged["searched"] = folders
    merged["candidates"] = candidates
    merged["note"] = None

    if not candidates:
        merged["path"], merged["confident"] = None, False
        merged["error"] = results[0]["error"] or "No file matched. Pick one below."
        return merged

    best = candidates[0]
    merged["folder"] = best["folder"]
    merged["path"] = best["path"]
    # Score 20 means the filename carried today's date, wherever it was found.
    merged["confident"] = best["score"] >= 20 or results[best["rank"]]["confident"]
    merged["error"] = None if merged["confident"] else results[best["rank"]]["error"]

    if best["rank"] > 0:
        passed_over = [c for c in candidates if c["rank"] == 0]
        if passed_over:
            merged["note"] = (
                f"Using the copy in {os.path.basename(best['folder'])}. "
                f"There is a match on the share ({passed_over[0]['name']}) but "
                f"its filename doesn't carry this date.")
    return merged


def _strip_extensions(name):
    """
    Take the extension off, however many times it was applied.

    Files get saved as "... 855028.06.xls.xls" often enough that one pass
    isn't enough, and a leftover ".xls" hides the amount at the end.
    """
    stem = name
    for _ in range(4):
        base, ext = os.path.splitext(stem)
        if ext.lower() not in (".xls", ".xlsx", ".xlsm", ".csv", ".pdf"):
            return stem
        stem = base
    return stem


def parse_check_filename(name):
    """
    Read the check count and dollar total straight off the filename. Seen so far:

        58 Check totaling $1,446,953.10.xls
        32 Check totaling $2,647,088.xls      (no cents)
        39 Checks 2338512.27.xls              (no "totaling", no $ or commas)
        29 Checks 855028.06.xls.xls           (extension applied twice)

    The count leads, the amount trails. Everything between them varies, so
    neither is matched on the wording around it. The file is never opened.
    """
    stem = _strip_extensions(name)
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
    # Fresh each time, so a file that appeared since the last look is seen.
    _listings.clear()
    umb = umb_folder()
    lock = lockbox_folder(d)
    dl = downloads_folder()

    sources = {
        # Both PDFs are looked for on the share and in Downloads. Which one
        # wins is decided by the filename score, not the folder order.
        "loans_tt": _find_in(
            [umb, dl], d,
            must_have=[("loan",), ("tt", "trade")],
            extensions=(".pdf",),
        ),
        "ach": _find_in(
            [umb, dl], d,
            must_have=[("ach",)],
            must_not_have=("loan",),
            extensions=(".pdf",),
        ),
        # Either pull will do. The 9:30 morning report and the 3:30 afternoon
        # one are the same workbook at two points in the day, so the sheet can
        # be built off the morning file and re-run later against the afternoon
        # one. The morning file arrives as .xlsx, the afternoon as .xls.
        "afternoon_report": _find(
            lock, d,
            must_have=[("ach",), ("afternoon", "morning")],
            extensions=(".xls", ".xlsx"),
            date_in_folder=True,
            prefer=("afternoon",),
        ),
        "check_total": _find(
            lock, d,
            # "Check totaling $X" and "Checks X" both turn up, so match on
            # "check" alone and keep the ACH reports out of the running.
            must_have=[("check",)],
            must_not_have=("ach", "afternoon", "morning"),
            extensions=(".xls", ".xlsx"),
            date_in_folder=True,
        ),
    }

    # The check file's name carries the UMB check count and dollar total.
    ct = sources["check_total"]
    ct["parsed"] = parse_check_filename(os.path.basename(ct["path"])) if ct["path"] else {}

    ar = sources["afternoon_report"]
    ar["parsed"] = (
        {"leading_count": parse_leading_count(os.path.basename(ar["path"])),
         "kind": afternoon.kind(ar["path"])}
        if ar["path"] else {}
    )
    # So the page can say "you're on the morning pull" without opening anything.
    for c in ar["candidates"]:
        c["kind"] = afternoon.kind(c["name"])

    return {"date": d.isoformat(), "base": BASE, "sources": sources,
            "output": output_path(d)}
