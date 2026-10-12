"""
Local server for the daily control sheet run.

    py server.py

Opens http://127.0.0.1:8777 in your browser. Stdlib only apart from
pdfplumber (and openpyxl once the writer lands).
"""
import json
import os
import shutil
import subprocess
import sys
import threading
import webbrowser
from datetime import date, datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs, unquote

import afternoon
import build
import discover
import items
import remittance
import upload_job

PORT = int(os.environ.get("CONTROL_SHEET_PORT", "8777"))
HERE = os.path.dirname(os.path.abspath(__file__))


def open_in_viewer(path):
    """Hand the file to whatever the OS uses for it."""
    if not os.path.exists(path):
        return False
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # noqa: S606 - intended behaviour
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
        return True
    except Exception:
        return False


def load_source(path, source, added=None):
    """Parse one Payment Center PDF and bucket its rows."""
    scan = items.scan(path)
    rows = scan["rows"]
    result = items.categorize(rows, source, added)
    return {
        "path": path,
        "row_count": len(rows),
        "integrity": {"ok": scan["ok"], "problems": scan["problems"],
                      "amounts_found": scan["amounts_found"]},
        "buckets": {
            name: {"rows": rows_, **items.totals(rows_)}
            for name, rows_ in result["buckets"].items()
        },
        "skipped": result["skipped"],
        "unmatched": result["unmatched"],
    }


def build_plan(body):
    """
    Assemble the workbook plan by re-reading the source files.

    The browser only supplies paths, the wire include/exclude choices, and the
    two figures that come off the emailed scan — so the workbook is always
    built from the files themselves, never from whatever the page happens to
    be holding.
    """
    d = datetime.strptime(body["date"], "%Y-%m-%d").date()

    # Rows dragged into a bucket by hand, as {payment_no: bucket}.
    added = body.get("added_rows") or {}
    tt = items.categorize(items.read_pdf(body["loans_tt"]), "loans_tt", added)
    ach = items.categorize(items.read_pdf(body["ach"]), "ach", added)
    report = afternoon.read(body["afternoon_report"])

    # Trade tickets, loan pulls and lockbox ACHs the user struck out. Sent as
    # payment numbers rather than positions, so they still point at the same
    # rows when this re-reads the file.
    dropped = set(body.get("dropped_rows") or [])
    tickets = items.keep("trade_tickets", tt["buckets"]["trade_tickets"], dropped)
    loans = items.keep("loan_pulls", tt["buckets"]["loan_pulls"], dropped)
    pulled = items.keep("ach_lockbox", ach["buckets"]["ach_lockbox"], dropped)

    # Wires the user chose to leave out, identified by their row in the report.
    dropped_wires = set(body.get("dropped_wires") or [])
    wires = [w["amount"] for w in report["wires"] if w["row"] not in dropped_wires]

    def entries(rows):
        return [{"name": r["recipient"], "amount": r["amount"]} for r in rows]

    return {
        "date": d,
        "loan_pulls": entries(loans),
        "trade_tickets": entries(tickets),
        "pm_ach": [e["amount"] for e in report["house"]],
        "wires": wires,
        "pulled_ach": [r["amount"] for r in pulled],
        "ach_statement_total": report["ach"]["total"],
        "ach_items": report["ach"]["items"],
        "remote_deposit_amount": body.get("remote_deposit_amount"),
        "remote_deposit_count": body.get("remote_deposit_count"),
        "umb_check_count": body.get("umb_check_count"),
        "lb2999_amount": body.get("lb2999_amount"),
    }


def lock_holder(path):
    """
    Who has this workbook open, if anyone. Best effort, and None means only
    "no sign of anyone" — never a guarantee that the file is free.

    Excel drops a hidden lock file named "~$<name>" beside a workbook while it
    is open, with the user's name inside: a length byte, then the name. Newer
    Excel writes it as UTF-16, older as plain bytes, so both are tried.

    Checking is worth it because the alternative is finding out from a failed
    copy halfway through, and because it can say *who* rather than just that
    something went wrong.
    """
    folder, name = os.path.split(path)
    lock = os.path.join(folder, "~$" + name)
    if not os.path.isfile(lock):
        return None
    try:
        with open(lock, "rb") as f:
            raw = f.read(256)
        count = raw[0] if raw else 0
        for text in (raw[1:1 + count * 2].decode("utf-16-le", "ignore"),
                     raw[1:1 + count].decode("latin-1", "ignore")):
            text = text.replace("\x00", "").strip()
            # A wrong guess at the encoding still decodes, just into nonsense,
            # so only accept something that reads like a name.
            if 2 <= len(text) <= 64 and all(
                    c.isalnum() or c in " .,'-_@" for c in text):
                return text
    except Exception:
        pass
    return "someone"


def run_generate(body):
    plan = build_plan(body)
    d = plan["date"]
    target = discover.output_path(d)
    overwrite = bool(body.get("overwrite"))

    scratch = os.path.join(HERE, "generated")
    os.makedirs(scratch, exist_ok=True)
    local = os.path.join(scratch, f"{d.isoformat()}-{d.day}.xlsx")
    build.generate(plan, local)

    result = {"target": target, "local": local,
              "download": f"/download/{os.path.basename(local)}"}

    # Someone already made today's sheet. Compare rather than write over it,
    # unless the overwrite box says otherwise — in which case still work out
    # the diff first, so what the overwrite changed is on the record.
    existing = target.get("exists")
    if existing:
        result["diff"] = build.diff(target["path"], plan)
        if not overwrite:
            result["status"] = "diff"
            result["message"] = ("A workbook already exists for today — comparing "
                                 "rather than overwriting.")
            return result

    if body.get("test_mode"):
        result["status"] = "test"
        result["message"] = "Test run — nothing written to the share."
        return result

    if target.get("error"):
        result["status"] = "no_target"
        result["message"] = target["error"] + " Download it and file it by hand."
        return result

    holder = lock_holder(target["path"])
    if holder:
        result["status"] = "locked"
        result["holder"] = holder
        result["message"] = (
            f"{target['path']} is open in Excel ({holder}), so it wasn't "
            f"written. Ask them to close it, then run this again — or download "
            f"the workbook below and file it by hand.")
        return result

    try:
        shutil.copyfile(local, target["path"])
    except PermissionError as exc:
        # Windows refuses the write while Excel holds the file. That is the
        # safe outcome: the alternative would be a half-copied workbook.
        result["status"] = "locked"
        result["message"] = (
            f"Windows wouldn't let us write to {target['path']} — it is most "
            f"likely open in Excel. Close it and run this again, or download "
            f"the workbook below. ({exc.strerror or exc})")
        return result

    result["status"] = "overwritten" if existing else "written"
    result["written_to"] = target["path"]
    result["message"] = (f"Overwrote {target['path']}" if existing
                         else f"Written to {target['path']}")
    return result


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *args):
        pass  # keep the console quiet

    # --- plumbing -------------------------------------------------------
    def _send(self, payload, status=200, content_type="application/json"):
        body = payload if isinstance(payload, bytes) else json.dumps(payload).encode()
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _body(self):
        length = int(self.headers.get("Content-Length", 0))
        return json.loads(self.rfile.read(length) or b"{}")

    # --- routes ---------------------------------------------------------
    def do_GET(self):
        route = urlparse(self.path)
        if route.path in ("/", "/index.html"):
            with open(os.path.join(HERE, "index.html"), "rb") as f:
                return self._send(f.read(), content_type="text/html; charset=utf-8")

        if route.path == "/api/locate":
            q = parse_qs(route.query)
            raw = q.get("date", [date.today().isoformat()])[0]
            try:
                d = datetime.strptime(raw, "%Y-%m-%d").date()
            except ValueError:
                return self._send({"error": f"Bad date: {raw}"}, 400)
            return self._send(discover.locate_all(d))

        # The userscripts poll this to decide whether to do anything at all.
        # No job armed means no automation running, which is the answer on any
        # day this isn't being used.
        if route.path == "/api/upload/job":
            return self._send(upload_job.status())

        if route.path == "/api/upload/next":
            q = parse_qs(route.query)
            return self._send(upload_job.next_pending((q.get("after") or [None])[0]))

        if route.path.startswith("/api/upload/for/"):
            return self._send(upload_job.for_deposit(route.path.rsplit("/", 1)[-1]))

        if route.path.startswith("/api/upload/file/"):
            got = upload_job.read(route.path.rsplit("/", 1)[-1])
            if not got:
                return self._send({"error": "No such file in the armed job"}, 404)
            name, mime, blob = got
            self.send_response(200)
            self.send_header("Content-Type", mime)
            self.send_header("Content-Disposition", f'attachment; filename="{name}"')
            self.send_header("X-Filename", name)
            self.send_header("Content-Length", str(len(blob)))
            self.end_headers()
            return self.wfile.write(blob)

        if route.path.startswith("/download/"):
            name = os.path.basename(unquote(route.path[len("/download/"):]))
            path = os.path.join(HERE, "generated", name)
            if not os.path.isfile(path):
                return self._send({"error": "Not found"}, 404)
            with open(path, "rb") as f:
                body = f.read()
            self.send_response(200)
            self.send_header(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition",
                             f'attachment; filename="{name}"')
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            return self.wfile.write(body)

        self._send({"error": "Not found"}, 404)

    def do_POST(self):
        route = urlparse(self.path).path
        try:
            body = self._body()
        except json.JSONDecodeError:
            return self._send({"error": "Malformed request body"}, 400)

        if route == "/api/open":
            paths = [p for p in body.get("paths", []) if p]
            opened = [p for p in paths if open_in_viewer(p)]
            failed = [p for p in paths if p not in opened]
            return self._send({"opened": opened, "failed": failed})

        if route == "/api/parse":
            out, errors = {}, {}
            for source in ("loans_tt", "ach"):
                path = body.get(source)
                if not path:
                    continue
                try:
                    out[source] = load_source(path, source, body.get("added_rows"))
                except Exception as exc:
                    errors[source] = str(exc)

            report = None
            if body.get("afternoon_report"):
                try:
                    report = afternoon.read(body["afternoon_report"])
                except Exception as exc:
                    errors["afternoon_report"] = str(exc)

            return self._send({"sources": out, "afternoon": report, "errors": errors})

        if route == "/api/checkname":
            # The count and amount live in the filename, so a file chosen by
            # hand has to be re-read the same way an auto-detected one is.
            path = body.get("path") or ""
            return self._send(discover.parse_check_filename(os.path.basename(path))
                              if path else {"count": None, "amount": None})

        if route == "/api/remittance":
            try:
                d = datetime.strptime(body["date"], "%Y-%m-%d").date()
                rows = items.categorize(
                    items.read_pdf(body["ach"]), "ach",
                    body.get("added_rows"))["buckets"]["ach_lockbox"]
                # A struck-out ACH isn't going on the sheet, so it shouldn't be
                # claiming a remittance file either.
                rows = items.keep("ach_lockbox", rows, body.get("dropped_rows"))
                return self._send(remittance.match(rows, d))
            except Exception as exc:
                return self._send({"error": str(exc), "rows": []})

        if route == "/api/upload/arm":
            try:
                d = datetime.strptime(body["date"], "%Y-%m-%d").date()
                rows = items.categorize(
                    items.read_pdf(body["ach"]), "ach",
                    body.get("added_rows"))["buckets"]["ach_lockbox"]
                rows = items.keep("ach_lockbox", rows, body.get("dropped_rows"))
                # Re-matched here rather than taken from the page, so the files
                # about to be uploaded are the ones on disk right now.
                found = remittance.match(rows, d)
                # match() sorts by amount; the split is pasted in PDF order and
                # the new deposits come back in that same order, so the ACHs are
                # re-laid in PDF order here and the files follow them.
                by_no = {r.get("payment_no"): r for r in found["rows"]}
                aches, files = [], []
                for r in rows:
                    m = (by_no.get(r["payment_no"]) or {}).get("match")
                    entry = {"payment_no": r["payment_no"], "amount": r["amount"],
                             "recipient": r["recipient"], "file_index": None,
                             "file": m["name"] if m else None}
                    if m:
                        entry["file_index"] = len(files)
                        files.append({"name": m["name"], "path": m["path"],
                                      "amount": r["amount"], "recipient": r["recipient"]})
                    aches.append(entry)
                unmatched = [r["recipient"] for r in found["rows"] if not r.get("match")]
                # The wires and in-house ACHs, so the deposit list can be
                # read without asking the browser to work anything out.
                figures = {"wires_in": [], "wires_out": [], "house": []}
                if body.get("afternoon_report"):
                    rep = afternoon.read(body["afternoon_report"])
                    out = set(body.get("dropped_wires") or [])
                    for w in rep["wires"]:
                        side = "wires_out" if (w["row"] in out or not w["include"]) \
                               else "wires_in"
                        figures[side].append({"amount": w["amount"],
                                              "sender": w["sender"]})
                    figures["house"] = [{"amount": e["amount"], "name": e["name"]}
                                             for e in rep["house"]]

                job = upload_job.arm(
                    deposit_date=d.isoformat(),
                    report_name=os.path.basename(body.get("report_name") or ""),
                    # Always the afternoon pull, whichever report built the
                    # sheet — the morning file is never the one uploaded here.
                    report_pattern="achs afternoon report",
                    figures=figures,
                    aches=aches,
                    files=files)
                job["unmatched"] = unmatched
                return self._send(job)
            except Exception as exc:
                return self._send({"armed": False, "error": str(exc)}, 200)

        if route == "/api/upload/deposits":
            return self._send(upload_job.set_deposits(body.get("deposits") or []))

        if route == "/api/upload/split/approve":
            return self._send(upload_job.approve_split(body.get("deposit_id")))

        if route == "/api/upload/split/report":
            return self._send(upload_job.report_split(body.get("status"),
                                                      body.get("message")))

        if route == "/api/upload/run":
            return self._send(upload_job.set_run(body.get("active"), body.get("reason")))

        if route == "/api/upload/step":
            return self._send(upload_job.record_step(
                body.get("action"), body.get("deposit_id"),
                body.get("ok"), body.get("message")))

        if route == "/api/upload/approve":
            return self._send(upload_job.approve(body.get("ids") or [],
                                                 body.get("auto_submit")))

        if route == "/api/upload/report":
            return self._send(upload_job.report(body.get("deposit_id"),
                                                body.get("status"), body.get("message")))

        if route == "/api/upload/session":
            return self._send(upload_job.set_session(body.get("session_id")))

        if route == "/api/upload/uploaded":
            return self._send(upload_job.mark_uploaded(body.get("indexes") or []))

        if route == "/api/upload/disarm":
            upload_job.disarm()
            return self._send({"armed": False})

        if route == "/api/generate":
            try:
                return self._send(run_generate(body))
            except build.Overflow as exc:
                return self._send({"status": "overflow", "message": str(exc)}, 200)
            except Exception as exc:
                return self._send({"status": "error", "message": str(exc)}, 200)

        self._send({"error": "Not found"}, 404)


def main():
    # Threaded, so one slow request can't freeze the rest.
    #
    # The share lives behind a VPN from home, and a folder read that takes ten
    # seconds would, on a single-threaded server, also block the page itself,
    # every other API call, and the download link — so the whole app looks
    # hung rather than one step being slow.
    server = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    server.daemon_threads = True
    url = f"http://127.0.0.1:{PORT}"
    print(f"Control sheet running at {url}")
    print(f"Source share: {discover.BASE}")
    print("Ctrl+C to stop.")
    threading.Timer(0.5, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")


if __name__ == "__main__":
    main()
