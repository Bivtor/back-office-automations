"""
Find the remittance file for each lockbox ACH.

Files live in a day folder like:
    Q:\\CORE\\PURCHASE TRADING\\ACH Remittance Detail\\ACH-2026\\08 - ACH\\06

and are named "<payer> <m-d> <amount>.xlsx", give or take:

    Sample Law LLC 8-5 455.39.xlsx
    Sample Resource Commission - ERISA 8-5 6215.31.xlsx
    Sample Christian School 8-5 $1,375.50 (2).xlsx

The amount is always last, sometimes carries a dollar sign or thousands
separators, and may be followed by "(n)" when one employer splits into several
files at the same amount.

Matching leans on the amount, which is exact, and uses the payer name only to
choose between files that share one. Names can't be trusted on their own: the
PDF truncates them ("Sample Resourc ...") and sometimes misspells them.
"""
import os
import re
from difflib import SequenceMatcher

from discover import BASE

# "<name> <m-d> <amount>" with an optional $ , and an optional (n) copy marker.
FILENAME_RE = re.compile(
    r"^(?P<name>.+?)\s+"
    r"(?P<date>\d{1,2}[-/]\d{1,2})\s+"
    r"\$?\s*(?P<amount>[\d,]+(?:\.\d{1,2})?)"
    r"\s*(?:\((?P<copy>\d+)\))?\s*$"
)

# Fallback when the date is missing: just take the trailing amount.
TRAILING_AMOUNT_RE = re.compile(
    r"^(?P<name>.+?)\s+\$?\s*(?P<amount>[\d,]+\.\d{2})"
    r"\s*(?:\((?P<copy>\d+)\))?\s*$"
)

EXTENSIONS = (".xlsx", ".xls", ".csv", ".pdf")


def folder(d):
    """The day's remittance folder, matched loosely so naming drift survives."""
    root = os.path.join(BASE, "ACH Remittance Detail")
    result = {"folder": None, "error": None}
    if not os.path.isdir(root):
        result["error"] = f"Folder not found: {root}"
        return result

    year_dir = next(
        (os.path.join(root, e) for e in sorted(os.listdir(root))
         if str(d.year) in e and os.path.isdir(os.path.join(root, e))), None)
    if not year_dir:
        result["error"] = f"No folder for {d.year} under {root}"
        return result

    month_dir = next(
        (os.path.join(year_dir, e) for e in sorted(os.listdir(year_dir))
         if os.path.isdir(os.path.join(year_dir, e))
         and re.match(rf"0?{d.month}\b", e.strip())), None)
    if not month_dir:
        result["error"] = f"No month folder for {d.month:02d} under {year_dir}"
        return result

    day_dir = next(
        (os.path.join(month_dir, e) for e in sorted(os.listdir(month_dir))
         if os.path.isdir(os.path.join(month_dir, e))
         and e.strip().lstrip("0") == str(d.day)), None)
    if not day_dir:
        result["error"] = f"No day folder for {d.day} under {month_dir}"
        return result

    result["folder"] = day_dir
    return result


def _amount(text):
    try:
        return round(float(str(text).replace(",", "").replace("$", "")), 2)
    except (TypeError, ValueError):
        return None


def _normalise(name):
    """Strip punctuation and case so 'Sample Law, LLC' matches 'Sample Law LLC'."""
    name = re.sub(r"\.\.\.|…", " ", str(name or ""))
    name = re.sub(r"[^A-Za-z0-9 ]+", " ", name)
    return re.sub(r"\s+", " ", name).strip().upper()


def _similarity(pdf_name, file_name):
    """
    Compare on the part of the name the PDF actually gave us.

    The recipient column truncates, so only its opening characters are
    meaningful — comparing full strings would punish a correct match for the
    words the PDF cut off.
    """
    a, b = _normalise(pdf_name), _normalise(file_name)
    if not a or not b:
        return 0.0
    if len(a) < len(b):
        b = b[:len(a)]
    return SequenceMatcher(None, a, b).ratio()


def read_folder(path):
    """Parse every file in the folder into name, amount, and copy number."""
    files = []
    for entry in sorted(os.listdir(path)):
        full = os.path.join(path, entry)
        if not os.path.isfile(full) or not entry.lower().endswith(EXTENSIONS):
            continue
        stem = os.path.splitext(entry)[0]
        m = FILENAME_RE.match(stem) or TRAILING_AMOUNT_RE.match(stem)
        payer = (m.group("name").strip() if m else stem)
        amount = _amount(m.group("amount")) if m else None
        copy = int(m.group("copy")) if m and m.group("copy") else None
        files.append({
            "name": entry,
            "path": full,
            "payer": payer,
            "amount": amount,
            "copy": copy,
            "parsed": bool(m),
            # Same payer, amount and copy number saved twice — "TSA 8-6
            # 192299.35" exists as both .pdf and .xls. Those are one
            # remittance in two formats, not two separate ones.
            "group": (_normalise(payer), amount, copy),
            "ext": os.path.splitext(entry)[1].lower(),
        })
    return files


def match(entries, day):
    """
    Pair each ACH row with a remittance file.

    Every file is claimed at most once, so when an employer splits into several
    files at the same amount each row gets its own rather than all of them
    pointing at the first.
    """
    located = folder(day)
    if located["error"]:
        return {"folder": None, "error": located["error"], "rows": [], "unused": []}

    files = read_folder(located["folder"])
    unparsed = [f for f in files if not f["parsed"]]

    # Score every row against every file, then hand out the best pairs first.
    scored = []
    for i, row in enumerate(entries):
        for f in files:
            amount_ok = f["amount"] is not None and _amount(row["amount"]) == f["amount"]
            sim = _similarity(row["recipient"], f["payer"])
            if not amount_ok and sim < 0.75:
                continue
            score = (70 if amount_ok else 0) + round(sim * 30)
            scored.append((score, i, f))

    # Ties go to the unsuffixed file first, then (1), then (2) — otherwise
    # "137.50 (1)" sorts ahead of "137.50" and the base copy is left over.
    scored.sort(key=lambda x: (-x[0], x[1], x[2]["copy"] or 0, x[2]["name"]))
    siblings = {}
    for f in files:
        siblings.setdefault(f["group"], []).append(f)

    taken_rows, taken_groups = {}, set()
    for score, i, f in scored:
        if i in taken_rows or f["group"] in taken_groups:
            continue
        others = [g["name"] for g in siblings[f["group"]] if g["path"] != f["path"]]
        taken_rows[i] = {"score": score, "also": others, **f}
        taken_groups.add(f["group"])

    claimed = {f["path"] for f in files if f["group"] in taken_groups}

    rows = []
    for i, row in enumerate(entries):
        pick = taken_rows.get(i)
        rows.append({
            "payment_no": row.get("payment_no"),
            "recipient": row.get("recipient"),
            "amount": row.get("amount"),
            "match": pick,
            "confidence": pick["score"] if pick else 0,
        })

    # Listed the way they get picked: largest amount first, but where one
    # employer splits into several files at the same amount, the copies run
    # in order — no suffix, then (1), then (2).
    rows.sort(key=lambda r: (
        -(_amount(r["amount"]) or 0),
        _normalise(r["match"]["payer"]) if r["match"] else _normalise(r["recipient"]),
        (r["match"]["copy"] or 0) if r["match"] else 0,
    ))

    return {
        "folder": located["folder"],
        "error": None,
        "rows": rows,
        "file_count": len(files),
        "unparsed": [f["name"] for f in unparsed],
        # Everything in the folder that no ACH row claimed. Plenty of files
        # here are uploaded through a different route, so a long list is
        # normal rather than a sign something went wrong.
        "other_files": [f["name"] for f in files if f["path"] not in claimed],
    }
