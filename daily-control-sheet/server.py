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
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import urlparse, parse_qs, unquote

import afternoon
import build
import discover
import items
import remittance

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


def load_source(path, source):
    """Parse one Payment Center PDF and bucket its rows."""
    scan = items.scan(path)
    rows = scan["rows"]
    result = items.categorize(rows, source)
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

    tt = items.categorize(items.read_pdf(body["loans_tt"]), "loans_tt")
    ach = items.categorize(items.read_pdf(body["ach"]), "ach")
    report = afternoon.read(body["afternoon_report"])

    # Wires the user chose to leave out, identified by their row in the report.
    dropped = set(body.get("dropped_wires") or [])
    wires = [w["amount"] for w in report["wires"] if w["row"] not in dropped]

    def entries(rows):
        return [{"name": r["recipient"], "amount": r["amount"]} for r in rows]

    return {
        "date": d,
        "loan_pulls": entries(tt["buckets"]["loan_pulls"]),
        "trade_tickets": entries(tt["buckets"]["trade_tickets"]),
        "pm_ach": [e["amount"] for e in report["house"]],
        "wires": wires,
        "pulled_ach": [r["amount"] for r in ach["buckets"]["ach_lockbox"]],
        "ach_statement_total": report["ach"]["total"],
        "ach_items": report["ach"]["items"],
        "remote_deposit_amount": body.get("remote_deposit_amount"),
        "remote_deposit_count": body.get("remote_deposit_count"),
        "umb_check_count": body.get("umb_check_count"),
        "lb2999_amount": body.get("lb2999_amount"),
    }


def run_generate(body):
    plan = build_plan(body)
    d = plan["date"]
    target = discover.output_path(d)

    scratch = os.path.join(HERE, "generated")
    os.makedirs(scratch, exist_ok=True)
    local = os.path.join(scratch, f"{d.isoformat()}-{d.day}.xlsx")
    build.generate(plan, local)

    result = {"target": target, "local": local,
              "download": f"/download/{os.path.basename(local)}"}

    # Someone already made today's sheet — compare instead of writing over it.
    if target.get("exists"):
        result["status"] = "diff"
        result["diff"] = build.diff(target["path"], plan)
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

    shutil.copyfile(local, target["path"])
    result["status"] = "written"
    result["written_to"] = target["path"]
    result["message"] = f"Written to {target['path']}"
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
                    out[source] = load_source(path, source)
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
                    items.read_pdf(body["ach"]), "ach")["buckets"]["ach_lockbox"]
                return self._send(remittance.match(rows, d))
            except Exception as exc:
                return self._send({"error": str(exc), "rows": []})

        if route == "/api/generate":
            try:
                return self._send(run_generate(body))
            except build.Overflow as exc:
                return self._send({"status": "overflow", "message": str(exc)}, 200)
            except Exception as exc:
                return self._send({"status": "error", "message": str(exc)}, 200)

        self._send({"error": "Not found"}, 404)


def main():
    server = HTTPServer(("127.0.0.1", PORT), Handler)
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
