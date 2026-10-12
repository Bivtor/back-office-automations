# Outstanding Rollover Report — GUI edition

Local web GUI wrapping `run.ps1`. No admin, no install, no system changes.
Same architecture as the Daily Queue Processor: a PowerShell `HttpListener`
owns all state, the work runs in a `Start-Job`, and the browser polls
`/status` once a second.

## Folder contents

Everything must sit in one folder:

```
START.bat        double-click this
server.ps1       HTTP listener, state, job lifecycle, path guessing
index.html       the GUI
run.ps1          the report generator (now with -JobMode)
config.json      redCompanies / stateTRS / internals / colors
```

`config.json` is the one you already updated. If it's missing, START.bat warns,
the GUI shows a red banner, and the run fails fast instead of silently
classifying every firm as blue.

## Using it

1. Download the Custom Rollover Search Results export as usual.
2. Double-click **START.bat**. A console window opens and the browser opens to
   `http://localhost:8080/`.
3. Both path fields are pre-filled by the same guessing `generate_command.ps1`
   did: the raw export by its known filename, and the newest file matching
   *outstanding rollover*. Each field also has a dropdown of the 12 most recent
   spreadsheets in Downloads.
4. **GENERATE REPORT**. Progress, phase, and the live log stream into the page.
5. When it finishes you get three panels: the sheet row counts, the **Pershing
   NetX rows to check by hand**, and the **near-miss review list**.
6. When it completes, the output folder opens automatically and the server
   stops. The results, Pershing and near-miss panels stay on screen — they're
   just frozen. Work the Pershing rows, then close the tab and the console
   window.

The console window is the server. Closing it stops everything. **"Shut down
server when finished" is on by default**; the folder is opened before the
server exits, because every button on the page is a call to the server and
they all go dead once it stops. Uncheck it if you want to keep the GUI live —
to run a second report without relaunching, or to keep "Open output folder"
working.

The workbook is named `MM.DD.YYYY Outstanding Rollover Report.xlsx`, inside a
folder stamped with the same date plus the run time.

## What changed in run.ps1

Only additive; running it by hand still behaves exactly as before.

- **`-JobMode` switch.** Set only by the server. Routes `Write-Log` through
  `Write-Output` so the server can drain it, suppresses every `Read-Host`, and
  emits the marker lines below.
- **`Read-OldReportSheetData` now fills a caller-supplied `-Into` hashtable**
  instead of returning one. It has to: the function logs *and* returned a
  value, so under `Write-Output` logging its own log lines would have been
  appended to its return value and the caller would have received an array
  instead of the hashtable. Every carried-over note and callback date would
  have been dropped silently. Taking the destination by reference removes the
  whole class of problem.
- **DEBUG never reaches the job stream.** Partly noise control (the red/blue
  match dump is hundreds of lines and would flush the server's 500-line
  buffer), but also correctness: `Test-FuzzyMatchAgainstList` returns a value
  and logs at DEBUG, so DEBUG output would corrupt every match result the same
  way. Everything still lands in the log file.

## Marker protocol

`run.ps1` writes these to stdout in JobMode; `server.ps1`'s `Parse-JobLine`
consumes them. Anything else is treated as a plain log line.

| Marker | Payload | Effect in the GUI |
|---|---|---|
| `@@STATUS` | `phase=N pct=N msg=text` | progress bar + phase line |
| `@@COUNT` | `<label>\|<rows>` | one row of the Results table |
| `@@PERSH` | free text | one row in the manual NetX panel |
| `@@NEARMISS` | free text | one row in the review panel |
| `@@DONE` | `workbook=.. folder=.. log=..` | status → Done, enables Open Folder |
| `@@FAIL` | `msg=text` | status → Failed |

The `key=value` parser uses a lookahead so values can contain spaces — needed
because the output folder name is `08.03.2026 Outstanding RO Report Generation
Request 17-30`. Verified round-trip on real paths.

## Endpoints

| Method | Path | Purpose |
|---|---|---|
| GET | `/` | serves index.html |
| GET | `/status` | drains the job, returns full state (polled every 1s) |
| GET | `/defaults` | guessed paths, recent files, staleness dates, config check |
| POST | `/start` | validates both paths, launches the job |
| POST | `/cancel` | stops the job |
| POST | `/openFolder` | Explorer at the output folder |
| POST | `/openWorkbook` | Explorer with the workbook selected |
| POST | `/shutdown` | stops the listener |

## Notes for distribution

- Port **8080**. If a colleague already has something on it, launch with
  `powershell -File server.ps1 -Port 8090` and browse there. The Daily Queue
  Processor also uses 8080, so don't run both at once.
- Requires the **ImportExcel** module, same as before. If it's missing the run
  fails immediately with a clear message rather than a silent instant
  "completion".
- The output folder is still created next to the raw input file, exactly as the
  console version did.
- Cancelling mid-run can leave a partially written workbook in the output
  folder. The GUI says so in the confirm dialog and in the log.

## Not yet done

`generate_command.ps1` opened the SOP document in the browser at startup. That
isn't wired in — say the word and it becomes a link in the GUI header, which is
less intrusive than a tab opening on every launch.
