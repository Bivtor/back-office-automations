# Back-Office Automations

A collection of small internal automation tools I built while working in the operations back office of a retirement-plan services firm: local-first utilities (PowerShell / Python / vanilla JS) that wrap manual workflows in a one-click launcher plus a local web GUI, and a set of Tampermonkey userscripts that smooth over the internal web apps those workflows run through.

Each folder is a standalone project. The common pattern for the desktop tools: a `start.bat` launcher, a local HTTP server (`server.ps1` / `server.py`) serving a single-page GUI, and a `run.ps1` / worker script that does the heavy lifting. No installs, no admin rights, stdlib-first. Several projects also ship a userscript that bridges the browser side of the workflow to the local server.

## Projects

| Project | What it does | Stack |
|---|---|---|
| [`auto-dq`](auto-dq/) | Daily queue processor. Parses an HTML-table Excel export, auto-computes work assignments for unassigned rollover rows, and produces styled audit ("original") and working ("modified") workbooks. Two userscripts close the loop: a scanner that reads product types from the internal web app and posts them back to the local server, and an assignment assistant that applies the "modified" sheet to the app's assignment modal with a color-coded preview and conflict blocking. | PowerShell, JS (userscripts), local web GUI |
| [`daily-control-sheet`](daily-control-sheet/) | Builds the daily cash-control sheet: parses payment-center PDFs, reconciles ACH totals, wires, and remittance files, and assembles the output workbook behind a local web GUI. A four-stage userscript pipeline then carries the remittance files into the deposit system: find the session, read and plan the deposit list, attach each file, and split the lockbox ACH. | Python (pdfplumber, stdlib HTTP server), HTML/JS, userscripts |
| [`outstanding-rollover-report`](outstanding-rollover-report/) | Generates the outstanding rollover report from a raw export with fuzzy source-company matching, carried-over notes and callback dates, results dashboard with live progress cards, plus a config editor GUI. | PowerShell, local web GUI |
| [`weekly-rollover-report`](weekly-rollover-report/) | Command generator and processor for the weekly rollover (transfers-in) report, driven by a JSON config for column order, sheet filters, and business-day holiday logic. | PowerShell |
| [`tsa-epars-remittance`](tsa-epars-remittance/) | TPA remittance processing: matches downloaded remittance zips/PDFs to expected targets by name mapping (`EPARS_NAMES.json`), unlocks password-protected PDFs, and organizes them into the day's ACH folder. | PowerShell |
| [`tcg-file-renamer`](tcg-file-renamer/) | Guided renamer that validates and standardizes TPA file names from an Excel manifest before they land on internal FTP drop folders. | PowerShell |
| [`task-timing-dashboard`](task-timing-dashboard/) | Single-file, fully offline task-timing dashboard: loads the team's timing workbook (SheetJS) and an optional SimpleAlloc CSV, renders D3 visualizations, and caches data locally. | Single HTML file, D3, SheetJS |
| [`userscripts`](userscripts/) | Twenty standalone Tampermonkey scripts for the internal apps (OES case system, SimpleAlloc deposit system, the bank's ACH portal, local PDFs): a responsive-layout rewrite of a fixed-pixel ExtJS app, a PDF-parsing trade-ticket validator, a bank-portal template finder that verifies routing and account numbers before filling an amount, and a dozen smaller click-savers. See the folder README for the full list. | JS (Tampermonkey), pdf.js, pdf-lib |
| [`misc-scripts`](misc-scripts/) | Small utilities: CSV shape checker (delimiter-count validation) and a filename dollar-amount normalizer. | PowerShell |

## Notes

- These are cleaned-up copies for reference; they were originally deployed as folder drops on shared drives and as Tampermonkey installs, so paths and config values reflect that environment.
- The employer is not named anywhere in this repo. Internal hostnames have been replaced with `example.com` placeholders (`oes.example.com`, `partners.example.com`), the employer's own name in data-matching config is `INTERNAL_COMPANY`, and the Windows profile in file paths is `USERNAME`. The userscript `@match` lines therefore will not fire as-is; point them at the real hosts to use them.
- Credentials have been redacted (`<REDACTED>`), coworker names and internal user IDs replaced with pseudonyms, client and employer names in config JSONs and code comments replaced with numbered or generic placeholders (`SOURCE_COMPANY_01`, `EMPLOYER_01`, `Sample Law LLC`).
- Third-party product names (the bank's payment portal, TPA firms, vendor libraries) are left as they are; they identify vendors, not the employer.
