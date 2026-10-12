"""
The job the upload userscripts work from.

The browser can't read a path off your disk and it can't be handed one, so the
files have to arrive as bytes. This module holds the one job that is currently
armed and serves those bytes, which keeps the whole queue on this side: the
userscripts stay stateless and re-entrant, so a mid-run crash means reloading a
page rather than working out how far it got.

Nothing is served until you arm a job from the control sheet, and an armed job
expires on its own. A local server that will hand out remittance files to
anything that asks, indefinitely, is not something to leave running.
"""
import os
import threading
import time
from collections import Counter

# How long an armed job stays live. Long enough to work through a day's
# deposits, short enough that a forgotten one isn't still serving files after
# lunch.
TTL_SECONDS = 45 * 60

MIME = {
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".xls":  "application/vnd.ms-excel",
    ".csv":  "text/csv",
    ".pdf":  "application/pdf",
}

_lock = threading.Lock()
_job = None


def _mime(path):
    return MIME.get(os.path.splitext(path)[1].lower(), "application/octet-stream")


# Wires from these senders come off the deposit list. Matched on the sender we
# read out of the report's ORG=/OBI= line, not on the amount, because the
# amount is shared and the sender is not.
DELETE_WIRE_SENDERS = ("BANK OF THE SIERRA", "NSCC")


def _cents(v):
    return None if v is None else int(round(float(v) * 100))


def _standard_pm(name):
    """A in-house ACH of the usual batch kind, "INTERNAL CORP,..."."""
    return (name or "").upper().replace(" ", "").startswith("INTERNALCORP")


def classify(deposits, figures, aches):
    """
    Work out what each row on the deposit list is, and what should happen to it.

    Matching is by amount, which is the only thing the list shows — no sender,
    no reference. That is fine for the wires, where the report gives a sender
    to check against, and it is not fine on its own for the split children,
    where the same figure can appear twice. Those are identified structurally
    instead: see split_children below.

    Returns every row with a verdict and an `action`, and nothing acts on it.
    """
    excluded, included, pm = {}, {}, {}
    for w in figures.get("wires_out") or []:
        excluded.setdefault(_cents(w["amount"]), []).append(w)
    for w in figures.get("wires_in") or []:
        included.setdefault(_cents(w["amount"]), []).append(w)
    for e in figures.get("house") or []:
        pm.setdefault(_cents(e["amount"]), []).append(e)

    # Amounts genuinely collide — three deposits at $900.00 and two at $150.00
    # on the day this was written — so a figure that matches one of those names
    # a group, not a row. Said out loud rather than papered over.
    seen = Counter(_cents(d.get("amount")) for d in deposits)

    out = []
    for d in deposits:
        c = _cents(d.get("amount"))
        verdict, why, who, action = "unknown", "Not one of the figures on the sheet.", None, "leave"
        if d.get("deleted"):
            verdict, action = "already_deleted", "leave"
            why = "Already struck out on the page."
        elif c in excluded:
            who = excluded[c][0].get("sender")
            verdict = "wire_left_out"
            listed = any(t in (who or "").upper() for t in DELETE_WIRE_SENDERS)
            action = "delete" if listed else "leave"
            why = (f"Wire left out of the sheet ({who})."
                   + (" On the delete list." if listed
                      else " Not on the delete list, so it stays."))
        elif c in included:
            who = included[c][0].get("sender")
            verdict, action = "wire_counted", "leave"
            why = f"Wire counted on the sheet ({who})."
        elif c in pm:
            who = pm[c][0].get("name")
            verdict, action = "house", "leave"
            why = "in-house ACH on the report."
        shared = seen[c] > 1 and verdict not in ("unknown", "already_deleted")
        if shared:
            why += (f" Careful: {seen[c]} deposits share this amount, so the "
                    f"figure doesn't say which of them is the one.")
        out.append({**d, "verdict": verdict, "why": why, "who": who,
                    "action": action, "ambiguous": shared})

    live = [d for d in out if not d.get("deleted")]

    # The split children, if the split has already been done.
    #
    # Identified by shape rather than by amount: a child is an ACH with no
    # split button of its own, and a split creates its children with
    # consecutive deposit IDs in the order the amounts were pasted. So sorting
    # by ID and looking for the ACHs as one unbroken run, in order, finds them
    # — and only the order tells two files at the same figure apart.
    #
    # Looked for as a run within the children rather than as the children
    # exactly, because the two failures that matter look different that way:
    # a stray child alongside a good split leaves exactly one run, while a
    # split done twice leaves the same run twice.
    want = [a["amount"] for a in aches or []]
    want_c = [_cents(x) for x in want]
    kids = sorted((d for d in live
                   if d.get("type", "").strip().upper() == "ACH"
                   and not d.get("can_split")),
                  key=lambda d: int(d["id"]) if str(d["id"]).isdigit() else 0)
    got = [d.get("amount") for d in kids]
    n = len(want)
    runs = [i for i in range(len(kids) - n + 1)
            if n and [_cents(k.get("amount")) for k in kids[i:i + n]] == want_c]
    matched = len(runs) == 1
    split_twice = len(runs) > 1
    ours = kids[runs[0]:runs[0] + n] if matched else []
    strays = [d for d in kids if d not in ours] if kids else []

    if matched:
        for d, a in zip(ours, aches):
            d["verdict"] = "split_child"
            d["ache"] = {k: a.get(k) for k in
                         ("payment_no", "amount", "recipient", "file", "file_index")}
            # The list turns a deposit's details button from yellow to grey
            # once a file is attached. Only "attached" skips it; "missing" and
            # "unknown" both still get the upload, because sending a file twice
            # is recoverable and silently skipping one is not noticed.
            if d.get("file") == "attached":
                d["action"] = "done"
                d["why"] = "Split from the in-house ACH, and already has a file attached."
            elif not a.get("file"):
                d["action"] = "no_file"
                d["why"] = ("Split from the in-house ACH, but no remittance file "
                            "matched it — this one needs doing by hand.")
            else:
                d["action"] = "upload"
                d["why"] = f"Split from the in-house ACH. Gets {a['file']}."
                if d.get("file") == "unknown":
                    d["why"] += (" Couldn't tell from the page whether it already "
                                 "has a file, so it's treated as not.")
        for d in strays:
            d["action"] = "review"
            d["why"] = ("An ACH with no split button that isn't one of today's "
                        "split children. Left alone — check what it is.")

    # Which in-house ACHs this is allowed to act on.
    #
    # The usual ones are the "INTERNAL CORP" batches. Anything else —
    # "INTERNAL SECUR, ISEC PYMT", "INTERNAL, ICO PYMT" — has turned out to
    # carry remittance of its own that someone else uploads, so it is held for
    # you rather than deleted, and is never the split.
    #
    # Deliberately not "hold anything with a file attached": the grey details
    # button that means "has a file" on a split child also shows on ordinary
    # in-house batches, and holding those would hold the deposit that needs
    # splitting and stall the whole run.
    pms = [d for d in live if d["verdict"] == "house"]
    for d in pms:
        if not _standard_pm(d.get("who")):
            d["action"] = "review"
            d["held"] = True
            d["why"] += (" Not touching it: its payee isn't the usual INTERNAL "
                         "CORP batch, and these have carried remittance of their "
                         "own. Delete it yourself if it should go.")
    usable = [d for d in pms if not d.get("held")]

    # The parent only comes off once its children are there to replace it.
    #
    # And a split is only ever proposed when there are no split children on
    # the list at all. If there are any and they don't line up, something has
    # gone wrong — a split done twice, a stray, a half-reversed merge — and the
    # one thing that can't help is splitting again. Whatever is largest at
    # that point isn't the deposit that was meant to be split: the real one
    # has usually been deleted already, so "the largest" is a smaller deposit
    # that was never supposed to be touched.
    split_target = None
    if usable and not kids:
        split_target = max(usable, key=lambda d: d.get("amount") or 0)
        tied = [d for d in usable
                if _cents(d.get("amount")) == _cents(split_target.get("amount"))]
        for d in usable:
            if d is split_target:
                d["action"] = "split"
                d["why"] += " Largest of them, so this is the one to split."
            else:
                d["action"] = "delete"
                d["why"] += " Not the one being split, so it comes off."
        if len(tied) > 1:
            for d in tied:
                d["ambiguous"] = True
                d["why"] += (f" {len(tied)} in-house ACHs are tied at this "
                             f"amount, so which to split isn't decidable here.")
    elif usable and matched:
        for d in usable:
            d["action"] = "delete"
            d["why"] += " Split is done, so the parent comes off."
    elif usable:
        # Children are on the list but they don't line up. Leave every
        # in-house ACH alone until that's sorted out by hand.
        for d in usable:
            d["action"] = "review"
            d["why"] += (" Held: the split children on the list don't line up, "
                         "so nothing is deleted or split until that's sorted out.")

    # The split page refuses to save unless the parent keeps something back:
    # its own calcRemaining() rejects a remainder of zero or less. So the ACHs
    # have to total strictly less than the deposit being split — the rare case
    # where they don't is worth seeing before the page throws it back.
    remainder_problem = None
    if split_target and want:
        room = (split_target.get("amount") or 0) - sum(want)
        if room <= 0:
            remainder_problem = (
                f"The {len(want)} ACHs total {sum(want):,.2f} but deposit "
                f"{split_target['id']} is only {split_target.get('amount') or 0:,.2f}. "
                f"The split page won't save unless something is left on the parent, "
                f"so this one needs doing by hand.")
            split_target["ambiguous"] = True
            split_target["why"] += " " + remainder_problem

    plan = {
        "split_done": matched,
        "split_twice": split_twice,
        "split_room": (None if not split_target
                       else round((split_target.get("amount") or 0) - sum(want), 2)),
        "split_target": split_target["id"] if split_target else None,
        "expected": want,
        "expected_sum": round(sum(want), 2),
        "found": got if want else [],
        "children": [d["id"] for d in ours],
        "held": [d["id"] for d in pms if d.get("held")],
        "problems": [],
    }
    if remainder_problem:
        plan["problems"].append(remainder_problem)
    if split_twice:
        plan["problems"].append(
            f"The {n} split amounts appear {len(runs)} times on the list, so the "
            f"split looks like it was done more than once. Merge the extra back "
            f"into the original before anything else runs — nothing will be split, "
            f"deleted or uploaded until there's only one set.")
    elif want and not matched:
        if not kids:
            plan["problems"].append(
                f"No split yet — nothing on the list looks like a child. "
                f"{len(want)} amounts are waiting to be split out.")
        else:
            plan["problems"].append(
                f"{len(kids)} deposit(s) look like split children but the "
                f"{len(want)} ACHs from the sheet aren't among them in order. "
                f"Not splitting again, and not pairing files on a partial match.")
    if strays and matched:
        plan["problems"].append(
            f"{len(strays)} ACH(s) with no split button aren't part of today's "
            f"split. Left alone — marked for review.")
    return out, plan


def arm(*, deposit_date, report_name, report_pattern, figures, aches, files):
    """
    Make a job live.

    `files` is [{name, path, ...}] — every remittance file the ACH rows
    matched, in the order section 8 listed them. Each gets an index that the
    userscripts fetch it by, so no path ever crosses into the browser and no
    path from the browser is ever opened.
    """
    global _job
    entries = []
    for i, f in enumerate(files):
        if not f.get("path") or not os.path.isfile(f["path"]):
            continue
        entries.append({
            "index": len(entries),
            "name": f.get("name") or os.path.basename(f["path"]),
            "path": f["path"],
            "type": _mime(f["path"]),
            "size": os.path.getsize(f["path"]),
            "amount": f.get("amount"),
            "recipient": f.get("recipient"),
            "url": None,           # filled in below, once the index is fixed
            "uploaded": False,
        })
    for e in entries:
        e["url"] = f"/api/upload/file/{e['index']}"

    with _lock:
        _job = {
            "armed_at": time.time(),
            "deposit_date": deposit_date,
            "report_name": report_name,
            # Stage 1 matches on this, not on the exact filename: the leading
            # item count changes every day and the sheet may have been built
            # from the morning pull, but the row to find is always the
            # afternoon one.
            "report_pattern": report_pattern,
            "figures": figures,
            "aches": aches,
            "deposits": [],
            "plan": {},
            # deposit id -> what was approved for it and how it went. Held here
            # so a reloaded tab can't send the same file a second time.
            "approved": {},
            # The delete run. Each delete reloads the list page, so the run
            # can't live in the page — it has to be picked up again from here
            # on the next load.
            "run": {"active": False, "stopped": None, "tries": {}, "log": []},
            "split": {"deposit_id": None, "status": None, "message": None},
            "files": entries,
            # Set by stage 1 once it has found the row, read by stage 2.
            "session_id": None,
            "missing": [f.get("name") for f in files
                        if not f.get("path") or not os.path.isfile(f["path"])],
        }
        return _public(_job)


def disarm():
    global _job
    with _lock:
        _job = None


def current():
    """The live job, or None once it has expired or was never armed."""
    global _job
    with _lock:
        if _job and time.time() - _job["armed_at"] > TTL_SECONDS:
            _job = None
        return _job


def status():
    job = current()
    return _public(job) if job else {"armed": False}


def _public(job):
    """What the userscripts get. Deliberately no filesystem paths."""
    return {
        "armed": True,
        "deposit_date": job["deposit_date"],
        "report_name": job["report_name"],
        "report_pattern": job["report_pattern"],
        "deposits": job["deposits"],
        "plan": job["plan"],
        "aches": job["aches"],
        "approved": job["approved"],
        "run": job["run"],
        "split": job["split"],
        "session_id": job["session_id"],
        "expires_in": int(TTL_SECONDS - (time.time() - job["armed_at"])),
        "missing": job["missing"],
        "files": [{k: v for k, v in f.items() if k != "path"} for f in job["files"]],
    }


def set_session(session_id):
    """Stage 1 tells us which upload session it landed on."""
    job = current()
    if not job:
        return {"armed": False}
    with _lock:
        job["session_id"] = str(session_id)
    return _public(job)


def mark_uploaded(indexes):
    job = current()
    if not job:
        return {"armed": False}
    with _lock:
        for i in indexes:
            if 0 <= i < len(job["files"]):
                job["files"][i]["uploaded"] = True
    return _public(job)


def set_deposits(deposits):
    """Stage 2 hands over the deposit list; we hand back what we make of it."""
    job = current()
    if not job:
        return {"armed": False}
    with _lock:
        job["deposits"], job["plan"] = classify(
            deposits, job["figures"], job["aches"])

        # The list is the site's own record of what each deposit holds, so it
        # settles any status we recorded along the way. A deposit whose button
        # has gone grey has its file, however it got there — including uploads
        # whose confirming page reload never got the chance to report back.
        for d in job["deposits"]:
            a = job["approved"].get(str(d["id"]))
            if a and d.get("file") == "attached" and a["status"] != "uploaded":
                a["status"] = "uploaded"
                a["message"] = "The deposit list shows a file attached."
                job["files"][a["file_index"]]["uploaded"] = True
    return _public(job)


def approve(ids, auto_submit=False):
    """
    Clear specific deposits for upload. This is the audit gate: stage 3 will
    not attach anything to a deposit that isn't in here.

    Only rows the plan currently marks "upload" can be approved, so a deposit
    whose button has since gone grey, or that was never a split child, is
    refused rather than sent a file.
    """
    job = current()
    if not job:
        return {"armed": False, "approved": [], "refused": list(ids)}
    ok, refused = [], []
    with _lock:
        rows = {str(d["id"]): d for d in job["deposits"]}
        for raw in ids:
            i = str(raw)
            d = rows.get(i)
            index = ((d or {}).get("ache") or {}).get("file_index")
            if not d or d.get("action") != "upload" or index is None:
                refused.append(i)
                continue
            job["approved"][i] = {"file_index": index, "amount": d.get("amount"),
                                  "auto_submit": bool(auto_submit),
                                  "status": "approved", "message": None,
                                  "at": time.time()}
            ok.append(i)
    return {"armed": True, "approved": ok, "refused": refused}


def next_pending(after=None):
    """
    The next approved deposit still waiting for its file.

    The queue is just the approved deposits in the order they were approved,
    minus the ones that are finished or have gone wrong. Kept here rather than
    in the browser so that a tab reloading itself — which is what the site does
    after every upload — doesn't lose its place.
    """
    job = current()
    if not job:
        return {"armed": False, "deposit_id": None}
    for i, a in job["approved"].items():
        if a["status"] in ("approved", "attached") and str(i) != str(after or ""):
            return {"armed": True, "deposit_id": i,
                    "left": sum(1 for x in job["approved"].values()
                                if x["status"] in ("approved", "attached"))}
    return {"armed": True, "deposit_id": None, "left": 0}


def for_deposit(deposit_id):
    """What stage 3 should put on this deposit, if anything."""
    job = current()
    if not job:
        return {"armed": False}
    a = job["approved"].get(str(deposit_id))
    if not a:
        return {"armed": True, "approved": False}
    f = job["files"][a["file_index"]]
    return {"armed": True, "approved": True, "status": a["status"],
            # What was recorded last time round, so a page that comes back to
            # a failure can say what the site actually replied rather than
            # only that the file isn't there.
            "message": a["message"],
            "auto_submit": a["auto_submit"], "expected_amount": a["amount"],
            "file": {k: v for k, v in f.items() if k != "path"}}


def report(deposit_id, status, message=None):
    """Stage 3 says how it went: attached, submitted, uploaded, or failed."""
    job = current()
    if not job:
        return {"armed": False}
    with _lock:
        a = job["approved"].get(str(deposit_id))
        if not a:
            return {"armed": True, "known": False}
        a["status"], a["message"] = status, message
        if status == "uploaded":
            job["files"][a["file_index"]]["uploaded"] = True
    return {"armed": True, "known": True, "status": status}


MAX_TRIES = 3


def set_run(active, reason=None):
    job = current()
    if not job:
        return {"armed": False}
    with _lock:
        job["run"]["active"] = bool(active)
        job["run"]["stopped"] = None if active else reason
        if active:
            job["run"]["tries"] = {}
            job["run"]["log"] = []
    return _public(job)


def record_step(action, deposit_id, ok=None, message=None):
    """
    Note an attempt, and stop the run if the same one keeps coming back.

    Each step is worked out afresh from the page after every reload, which is
    what keeps the run honest — but it also means a step that silently fails
    would be chosen again forever. Counting attempts turns that into a stop
    with a reason rather than a loop.
    """
    job = current()
    if not job:
        return {"armed": False, "stop": True}
    key = f"{action}:{deposit_id}"
    with _lock:
        run = job["run"]
        run["tries"][key] = run["tries"].get(key, 0) + 1
        tries = run["tries"][key]
        run["log"].append({"action": action, "id": str(deposit_id), "ok": ok,
                           "message": message, "try": tries, "at": time.time()})
        run["log"] = run["log"][-60:]
        stop = tries > MAX_TRIES
        if stop:
            run["active"] = False
            run["stopped"] = (f"Tried to {action} deposit {deposit_id} {tries} times "
                              f"and it is still on the list. Stopped.")
        return {"armed": True, "stop": stop, "tries": tries,
                "reason": run["stopped"], "run": run}


def approve_split(deposit_id):
    """
    Clear one deposit to be split. Only the deposit the plan actually named,
    so the split page can't be pointed at something else, and not at all when
    the amounts wouldn't leave a remainder.
    """
    job = current()
    if not job:
        return {"armed": False, "ok": False}
    target = (job["plan"] or {}).get("split_target")
    room = (job["plan"] or {}).get("split_room")
    if not target or str(target) != str(deposit_id):
        return {"armed": True, "ok": False,
                "message": f"The plan's split is deposit {target}, not {deposit_id}."}
    if room is not None and room <= 0:
        return {"armed": True, "ok": False,
                "message": "The ACHs don't leave a remainder on the parent."}
    with _lock:
        # One split per run, and once the amounts have gone into the form it
        # counts as done whether or not the page confirmed it. Re-approving
        # used to reset a saved split back to "approved", which is how a page
        # refresh could put the same split in twice. Only a split that was
        # refused or failed before anything was entered can be tried again.
        prior = job["split"]
        if prior.get("deposit_id") and prior.get("status") in ("filled", "submitted"):
            return {"armed": True, "ok": False,
                    "message": (f"A split was already entered in this run "
                                f"(deposit {prior['deposit_id']}). Not doing another. "
                                f"Re-arm from the control sheet if it really needs redoing.")}
        job["split"] = {"deposit_id": str(deposit_id), "status": "approved",
                        "message": None}
    return {"armed": True, "ok": True, "deposit_id": str(deposit_id)}


def report_split(status, message=None):
    job = current()
    if not job:
        return {"armed": False}
    with _lock:
        job["split"]["status"] = status
        job["split"]["message"] = message
    return _public(job)


def read(index):
    """
    The bytes of one file, by its index in the armed job.

    Indexed rather than named because an index can't be turned into a path.
    Anything the browser sends is only ever used to pick from the list this
    side built, so there is nothing to escape out of.
    """
    job = current()
    if not job:
        return None
    try:
        entry = job["files"][int(index)]
    except (ValueError, TypeError, IndexError):
        return None
    try:
        with open(entry["path"], "rb") as f:
            return entry["name"], entry["type"], f.read()
    except OSError:
        return None
