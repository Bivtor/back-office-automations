# Userscripts

Standalone Tampermonkey scripts for the internal web apps the back-office workflows run through. Each file is self-contained; install it in Tampermonkey and point the `@match` lines at the real hosts (they are `example.com` placeholders here).

Scripts that belong to a larger project live with that project instead:

- [`../auto-dq/persh_scanner.user.js`](../auto-dq/persh_scanner.user.js) and [`../auto-dq/assignment_assistant.user.js`](../auto-dq/assignment_assistant.user.js)
- [`../daily-control-sheet/userscripts/`](../daily-control-sheet/userscripts/) (the four-stage remittance upload pipeline)

## OES (case-management system, ExtJS + Bootstrap SPA)

| Script | What it does |
|---|---|
| [`responsive-layout`](oes/responsive-layout.user.js) | Replaces the app's fixed-pixel layout with flex/grid and breakpoints via a `@layer` stylesheet that beats the site's `!important` rules without specificity games. Clamps ExtJS / jQuery UI / jBox popups inside the viewport with Resize and Mutation observers, adds click-to-copy on every OES ID, and reopens the case you were in after an F5 by wrapping the site's own `showEdit` / `changeState` functions. Alt+Shift+R toggles it for comparison. |
| [`trade-ticket-ach-checker`](oes/trade-ticket-ach-checker.user.js) | Intercepts clicks on Trade Ticket PDFs in the Purchase Team queue, opens them in a shell window, extracts the text with pdf.js using row/column geometry, validates routing number, bank account type, annual contribution limit (by tax year and catch-up age) and the large-amount disclosure, then produces the one-line payload the UMB scripts consume. Copy is blocked until the checks pass. |
| [`better-uploader`](oes/better-uploader.user.js) | Auto-fills the display name from the chosen file, picks the document type from the filename, opens the file picker as soon as the modal appears, Enter uploads, and the case refreshes on success. |
| [`better-documents-list`](oes/better-documents-list.user.js) | Natural, case-insensitive sort of every "Current Documents" list (viewer, fax modal, any other instance) with an asc/desc caret. |
| [`doc-renamer-enter-fix`](oes/doc-renamer-enter-fix.user.js) | Makes Enter save an inline document rename instead of navigating away. |
| [`column-chooser-presets`](oes/column-chooser-presets.user.js) | One-click column sets for the queue's ColVis chooser: the columns the weekly rollover report and the outstanding rollover report expect. |
| [`copy-client-id`](oes/copy-client-id.user.js) | Click or Alt+hover the Client ID to copy the full identifier the UI only shows truncated. |
| [`internal-note-quick-fill`](oes/internal-note-quick-fill.user.js) | Canned-note buttons in the Add Internal Note modal, only on the Purchase Team tab. |
| [`msg-stamp-copier`](oes/msg-stamp-copier.user.js) | Copies a tab-separated row (date, client, type, expected amount, rep, source company, OES ID) for pasting into the tracking sheet. |
| [`tpa-form-flag`](oes/tpa-form-flag.user.js) | Flags TPAs that require a TPA form in the Employer Summary modal. |

## SimpleAlloc (deposit and rollover maintenance)

| Script | What it does |
|---|---|
| [`alt-click-date-fill`](simplealloc/alt-click-date-fill.user.js) | Alt+click any input to insert today's date. |
| [`tpa-warning-on-105s`](simplealloc/tpa-warning-on-105s.user.js) | Warns when a TPA is selected on a 105 rollover into a 403(b), where no TPA is needed. |
| [`copy-deposit-amount`](simplealloc/copy-deposit-amount.user.js) | Click the deposit total to copy a clean numeric value. |
| [`paste-splitter`](simplealloc/paste-splitter.user.js) | Paste a column of amounts from Excel into the Split ACH form; rows are added as needed. |
| [`rad-row-counter`](simplealloc/rad-row-counter.user.js) | Shows a running count of items processed through the RAD screen, frame-aware. |

## UMB (bank ACH portal, Angular)

These consume the payload produced by the trade-ticket checker above: `OES # … Name RN ######### AN ##### Checking $0.00`.

| Script | What it does |
|---|---|
| [`payment-center-quick-search`](umb/payment-center-quick-search.user.js) | One-click search presets ("today's trade tickets and loan pulls", "today's ACHs") with Federal Reserve holiday-aware next-business-day logic. |
| [`ach-template-find-and-verify`](umb/ach-template-find-and-verify.user.js) | Searches PPD templates by surname, disambiguates duplicates by the account's last four digits, opens the payment, verifies routing / account / account type against the ticket, and only then fills the amount. Nothing is persisted; the payload lives in the tab. |
| [`ach-template-fill-from-trade-ticket`](umb/ach-template-fill-from-trade-ticket.user.js) | Fills a new ACH PPD template and recipient form from the payload with simulated keystrokes so the Angular bindings see every change, with a step log for anything that needs a manual pick. |

## PDF helpers (`file://` pages)

| Script | What it does |
|---|---|
| [`auto-jb-stamp`](pdf/auto-jb-stamp.user.js) | For PDFs named `eq_<account>_<box>`, draws a red X on page 8 with pdf-lib and saves a `_stamped` copy, with a loop guard so the download doesn't retrigger itself. |
| [`auto-save-pdf-from-outlook`](pdf/auto-save-pdf-from-outlook.user.js) | Saves PDFs opened from the Outlook cache into Downloads with a timestamped filename. |

## Notes

- `@require` libraries (pdf.js 3.11.174, pdf-lib 1.17.1) load from public CDNs; they are not vendored here.
- Scripts that talk to a local server use `GM_xmlhttpRequest` with `@connect localhost` rather than `fetch`, so the server never has to send permissive CORS headers.
