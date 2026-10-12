// ==UserScript==
// @name         OES Responsive Layout
// @namespace    http://tampermonkey.net/
// @version      0.3
// @description  Replaces OES's fixed-pixel layout with flex/grid + breakpoints, keeps popups inside the viewport, click-to-copy OES IDs, and reopens your OES after a reload
// @author       Victor
// @match        https://oes.example.com/*
// @run-at       document-start
// @grant        GM_setClipboard
// ==/UserScript==
(function () {
    'use strict';

    /*
     * HOW THE OVERRIDES WIN
     * Everything lives in @layer oesfix and uses !important. An !important
     * declaration inside a cascade layer beats every unlayered !important
     * rule on the page (the site's inline <style> blocks, Bootstrap's w-25,
     * etc.) regardless of selector specificity, and !important stylesheet
     * rules also beat non-important inline style="" attributes. So we never
     * have to play specificity games with the site's CSS.
     *
     * The one thing we must NOT do is force `display` on anything the site
     * hides with inline display:none, hence the :not([style*="none"]) guards.
     *
     * Toggle the whole thing on/off with Alt+Shift+R to compare.
     */

    const CSS = String.raw`
@layer oesfix {

/* ============================================================
   0. Tokens
   ============================================================ */
:root {
    --oes-gap: 12px;
    --oes-edge: 8px;          /* min distance between popups and viewport edge */
    --oes-line: var(--bs-border-color, #dee2e6);
    --oes-bg: var(--bs-body-bg, #fff);
}

/* Page-level modal mask is sized in px at open time and never updates */
body > .x-mask {
    position: fixed !important;
    inset: 0 !important;
    width: auto !important;
    height: auto !important;
}

/* ============================================================
   1. Top nav
   ============================================================ */
/* The big "OES" banner image: pure decoration, costs ~45px of height */
#navHeader {
    display: none !important;
}

#nav {
    display: flex !important;
    flex-wrap: wrap !important;
    gap: 4px !important;
    width: auto !important;
    height: auto !important;
    padding: 6px 48px 6px 0 !important;   /* right side clears the floating search magnifier */
}
#nav .navBox {
    float: none !important;
    display: inline-flex !important;
    align-items: center !important;
    gap: 6px !important;
    width: auto !important;
    height: auto !important;
    margin: 0 !important;
    white-space: nowrap !important;
}
#nav .navBox img {
    max-height: 20px !important;
    width: auto !important;
    margin: 0 !important;
}

/* ============================================================
   2. Queue / search view  (#searchMain)
   ============================================================ */
#searchMain:not([style*="display: none"]):not([style*="display:none"]) {
    display: grid !important;
    grid-template-columns: minmax(0, 1fr) !important;
    gap: var(--oes-gap) !important;
    margin: 0 !important;
    overflow: visible !important;
}
#searchMain > [class*="col-"] {
    flex: none !important;
    width: auto !important;
    max-width: none !important;
    padding: 0 !important;
    min-width: 0 !important;
}
@media (min-width: 1100px) {
    #searchMain:not([style*="display: none"]):not([style*="display:none"]) {
        grid-template-columns: minmax(0, 1fr) clamp(280px, 24vw, 360px) !important;
        align-items: stretch !important;
    }
}
/* Search panel: fill the full height of the row so its border line
   runs all the way down beside the list instead of stopping halfway */
#newAcctSearchWin,
#newAcctSearchWin > .jbox-container {
    height: auto !important;
    min-height: 100% !important;
    box-sizing: border-box !important;
}
#newAcctSearchWin > .jbox-container {
    height: 100% !important;
}

/* Site forces overflow-x:scroll (permanent scrollbar) - make it auto.
   Height: grow with the list up to 80% of the screen, then scroll. */
#searchMain #searchResults,
#searchMain .queueSearchResults {
    height: auto !important;
    min-height: min(300px, 80vh) !important;
    max-height: 80vh !important;
    overflow-x: auto !important;
    overflow-y: auto !important;
}
#resultsTable { width: 100% !important; margin-bottom: 0 !important; }
#resultsTable th,
#resultsTable td { white-space: nowrap !important; }

/* Column header bar: give it some breathing room */
#searchResults table th,
#searchResults table tr.title > td {
    padding-top: 9px !important;
    padding-bottom: 9px !important;
    vertical-align: middle !important;
    line-height: 1.2 !important;
}

#newAcctSearchWin .jbox-container,
#newAcctSearchWin .jBox-content {
    width: auto !important;
    max-width: none !important;
}
#mainSearch .container {
    max-width: none !important;
    padding-inline: 0 !important;
}
#mainSearch .row {
    --bs-gutter-x: 8px;
    row-gap: 6px !important;
    margin-bottom: 6px !important;
}
#mainSearch select,
#mainSearch input.form-control {
    width: 100% !important;
    max-width: 100% !important;
}
#recentViewed > div[style*="497px"] {
    height: auto !important;
    max-height: min(60vh, 520px) !important;
}

/* The bordered box around the queue had a fixed height from style.css,
   so the taller list spilled out of it. Let it size to its contents. */
#searchMain {
    height: auto !important;
    min-height: 0 !important;
    max-height: none !important;
}

/* Assignment Admin (left) + Export to Excel (right): the site pinned
   these at a fixed spot, so they landed on top of the taller list.
   Put both back in normal flow under the box, sharing one line. */
#sectionLinks:not([style*="none"]) {
    position: static !important;
    inset: auto !important;
    float: none !important;
    display: flex !important;
    align-items: center !important;
    gap: 6px !important;
    width: auto !important;
    height: auto !important;
    min-height: 34px !important;
    margin: 8px 0 0 !important;
    padding: 0 180px 0 0 !important;   /* leave room for the Export button */
    z-index: auto !important;
}
#ExportToExcelBtn:not([style*="none"]) {
    position: static !important;
    inset: auto !important;
    float: none !important;
    display: flex !important;
    justify-content: flex-end !important;
    align-items: center !important;
    width: auto !important;
    height: 34px !important;
    margin: -34px 0 0 !important;       /* pull up onto the Assignment Admin line */
    z-index: auto !important;
    pointer-events: none !important;     /* the empty left side must not block the link */
}
#ExportToExcelBtn #exportBtn {
    pointer-events: auto !important;
    margin: 0 !important;
}

/* ============================================================
   3. Open OES: header strip (OES ID / Client ID / State / Rep)
   ============================================================ */
#editHeader .row {
    display: flex !important;
    flex-wrap: wrap !important;
    gap: 2px 16px !important;
    margin: 0 !important;
}
#editHeader .row > [class*="col-"] {
    flex: 1 1 auto !important;
    width: auto !important;
    max-width: none !important;
    padding: 0 !important;
}
@media (max-width: 899.98px) {
    #editHeader .text-center,
    #editHeader .text-end { text-align: left !important; }
}

/* ============================================================
   3b. Open OES fills exactly one screen - no page scroll.
       While an OES is open, the page is locked to the viewport
       height and the OES card takes all the space under the nav.
       If the content is taller than that, the OES body scrolls
       inside the card; Save / Close stay pinned at the bottom.
       Only applies while #mainEdit is showing an OES; the queue
       view scrolls normally.
   ============================================================ */
body:has(#mainEdit:not([style*="none"]) > .card) {
    height: 100dvh !important;
    overflow: hidden !important;
    display: flex !important;
    flex-direction: column !important;

    & > .container-fluid {
        flex: 1 1 auto !important;
        min-height: 0 !important;
        display: flex !important;
        flex-direction: column !important;
    }
    & > .container-fluid > div:has(> form[name="mainStateForm"]),
    & form[name="mainStateForm"] {
        flex: 1 1 auto !important;
        min-height: 0 !important;
        display: flex !important;
        flex-direction: column !important;
    }
    & #mainEdit {
        flex: 1 1 auto !important;
        min-height: 0 !important;
        display: flex !important;
        flex-direction: column !important;
    }
    & #mainEdit > .card {
        flex: 1 1 auto !important;
        min-height: 0 !important;
        display: flex !important;
        flex-direction: column !important;
        margin: 0 !important;
    }
    & #mainEdit > .card > .card-body {
        flex: 1 1 auto !important;
        min-height: 0 !important;
        overflow: auto !important;          /* fallback only - on tiny screens */
        display: flex !important;
        flex-direction: column !important;
    }
    & #mainEdit > .card > .card-body > .row {
        flex: 1 1 auto !important;
        min-height: 0 !important;
    }

    /* The three panes each scroll themselves, so the container doesn't.
       Roll Over + Documents are capped, Notes takes whatever's left. */
    & #mainEdit .card-body:has(> .roInfoTable) {
        max-height: 38vh !important;
        overflow: auto !important;
    }
    & #mainEdit #col3Info {
        max-height: 38vh !important;
        overflow: auto !important;
    }
    & #mainEdit > .card > .card-body > .row > :has(#userNotes) {
        display: flex !important;
        flex-direction: column !important;
        min-height: 0 !important;
    }
    & #mainEdit > .card > .card-body > .row > :has(#userNotes) > .card {
        flex: 1 1 auto !important;
        min-height: 0 !important;
        height: auto !important;
        display: flex !important;
        flex-direction: column !important;
    }
    & #mainEdit .card-body:has(> #userNotes) {
        flex: 1 1 auto !important;
        min-height: 0 !important;
        overflow: auto !important;
    }

    /* Transfers-style layout: top row = Roll Over | Documents,
       bottom row = Notes, stretched to fill */
    @media (min-width: 900px) {
        & #mainEdit:not(:has([onclick*="saveTTApp"])) > .card > .card-body > .row:has(> * #col3Info) {
            display: grid !important;
            grid-template-columns: minmax(0, 1fr) fit-content(460px) !important;
            grid-template-rows: auto minmax(0, 1fr) !important;
            align-items: stretch !important;

            & > * {
                width: auto !important;
                max-width: none !important;
                flex: none !important;
            }
            & > :has(#userNotes) { grid-column: 1 / -1 !important; }
        }
    }

    /* Trade Ticket grid: last row (Notes) stretches to fill */
    @media (min-width: 900px) and (max-width: 1399.98px) {
        & #mainEdit:has([onclick*="saveTTApp"]) > .card > .card-body > .row:has([name="docStatus"]):has(#col3Info):has(#userNotes) {
            grid-template-rows: auto auto minmax(0, 1fr) !important;
        }
    }
    @media (min-width: 1400px) {
        & #mainEdit:has([onclick*="saveTTApp"]) > .card > .card-body > .row:has([name="docStatus"]):has(#col3Info):has(#userNotes) {
            grid-template-rows: auto minmax(0, 1fr) !important;
        }
    }
    /* Footer is already at the bottom of the card now; no need to stick */
    & #editFooter {
        position: static !important;
        flex: 0 0 auto !important;
    }
}

/* ============================================================
   4. Open OES, Trade Ticket modules (Purchase Team etc.):
      Status / Client Info / Documents / Notes grid.
      Documents column is fit-content: as narrow as its longest
      row, capped at 460px. Client Info takes the rest.
   ============================================================ */
#mainEdit:has([onclick*="saveTTApp"]) .card-body > .row:has([name="docStatus"]):has(#col3Info):has(#userNotes) {
    display: grid !important;
    gap: var(--oes-gap) !important;
    margin: 0 !important;
    grid-template-columns: minmax(0, 1fr) !important;
    grid-template-areas: "status" "client" "docs" "notes" !important;

    & > * {
        flex: none !important;
        width: auto !important;
        max-width: none !important;
        min-width: 0 !important;
        padding: 0 !important;
        margin: 0 !important;
    }
    & > :has([name="docStatus"])                         { grid-area: status !important; }
    & > :has(#col3Info)                                  { grid-area: docs !important; min-width: min(300px, 100%) !important; }
    & > :has(#userNotes)                                 { grid-area: notes !important; }
    & > :not(:has([name="docStatus"], #col3Info, #userNotes)) { grid-area: client !important; }

    & > * > .card {
        height: 100% !important;
        margin: 0 !important;
    }

    /* the hard-coded 250px card bodies */
    & .card-body[style*="250px"] { height: auto !important; }
    & #col3Info { max-height: 360px !important; overflow: auto !important; }

    & .mt-5 { margin-top: 1rem !important; }
    & select[name="docStatus"],
    & select[name="assignedTo"] { width: 100% !important; }

    /* medium: Status + Documents side by side, Client Info full width */
    @media (min-width: 900px) {
        grid-template-columns: minmax(220px, 1fr) fit-content(460px) !important;
        grid-template-areas:
            "status docs"
            "client client"
            "notes  notes" !important;
    }
    /* wide: the original 3-across, but with sane minimums */
    @media (min-width: 1400px) {
        grid-template-columns: minmax(210px, 260px) minmax(0, 1fr) fit-content(460px) !important;
        grid-template-areas:
            "status client docs"
            "notes  notes  notes" !important;
    }
}

/* ============================================================
   4b. Open OES, everything else (Transfers / Roll Over etc.):
       Documents shrinks to its content (max 460px), every other
       card flex-grows into the leftover space. Full-width (-12)
       columns like Notes always get their own line.
   ============================================================ */
#mainEdit:not(:has([onclick*="saveTTApp"])) .card-body > .row:has(> * #col3Info) {
    display: flex !important;
    flex-wrap: wrap !important;
    align-items: stretch !important;
    gap: var(--oes-gap) !important;
    margin: 0 !important;

    & > * {
        padding: 0 !important;
        margin: 0 !important;
        min-width: 0 !important;
    }
    /* Roll Over Information (and any other side card): grow */
    & > :not(:has(#col3Info)):not([class*="-12"]) {
        flex: 1 1 0 !important;
        width: auto !important;
        max-width: none !important;
        min-width: min(340px, 100%) !important;
    }
    /* Documents: shrink-wrap, but never below 300px - on save the site
       empties and re-fills this list, and without a floor the column
       collapsed to nothing for a moment and Roll Over jumped across */
    & > :has(#col3Info) {
        flex: 0 1 auto !important;
        width: auto !important;
        max-width: min(460px, 100%) !important;
        min-width: min(300px, 100%) !important;
    }
    & > [class*="col-"][class*="-12"] { flex: 1 1 100% !important; }

    /* Each column is a flex column. The Roll Over column holds the
       Roll Over card PLUS a nested row (Document Status / Assignment),
       so the card grows to fill and the nested row keeps its height. */
    & > * {
        display: flex !important;
        flex-direction: column !important;
        gap: var(--oes-gap) !important;
    }
    & > * > .card {
        flex: 1 1 auto !important;
        height: auto !important;
        margin: 0 !important;
    }
    & > * > .row {
        flex: 0 0 auto !important;
        display: flex !important;
        flex-wrap: wrap !important;
        gap: var(--oes-gap) !important;
        margin: 0 !important;
    }
    & > * > .row > * {
        flex: 1 1 240px !important;
        width: auto !important;
        max-width: none !important;
        padding: 0 !important;
        margin: 0 !important;
    }
    & > * > .row > * > .card {
        height: 100% !important;
        margin: 0 !important;
    }
    & .card-body[style*="height"] { height: auto !important; }
    & #col3Info { max-height: 420px !important; overflow: auto !important; }

    @media (max-width: 899.98px) {
        & > :has(#col3Info) {
            flex: 1 1 100% !important;
            max-width: none !important;
        }
    }
}

/* Document Status: drop the useless header, and un-fix any
   hard-coded height on whatever wraps the dropdown so it can't
   spill out underneath Notes. Works in every module.
   Only the innermost card that holds the dropdown - the big outer
   OES card also "contains" it, and its header is the OES ID /
   Client / Rep strip, which must stay. */
#mainEdit .card:has([name="docStatus"]):not(:has(.card [name="docStatus"])) > .card-header { display: none !important; }
#mainEdit .card-body [style*="height"]:has([name="docStatus"]),
#mainEdit .card-body [style*="height"]:has([name="assignedTo"]) {
    height: auto !important;
    overflow: visible !important;
}

/* Table header bars inside an open OES (Roll Over columns,
   "Current Documents" / "Related Client Profiles" bars) */
#mainEdit table th,
#mainEdit .docTable td.currentDocs,
#mainEdit .docTable td.relatedDocs {
    padding-top: 9px !important;
    padding-bottom: 9px !important;
    vertical-align: middle !important;
    line-height: 1.2 !important;
}

/* Right-click rename box in the Documents list: exactly as tall as the
   line of text it replaces, so the row doesn't grow and jitter.
   The outline is drawn inside the box, so it adds no size either. */
.docTable input[name="loadDoc_editDisplayName"] {
    box-sizing: border-box !important;
    width: 100% !important;
    height: 1.5em !important;               /* fallback */
    height: 1lh !important;                 /* = the row's own line height */
    margin: 0 !important;
    padding: 0 3px !important;
    border: 0 !important;
    outline: 1px solid #6699cc !important;
    outline-offset: -1px !important;
    font: inherit !important;
    line-height: inherit !important;
    vertical-align: top !important;
    background: #fff !important;
}

/* Inner cards become size containers so their contents respond to the
   card's width, not the window's. NOT the Documents card: inline-size
   containment makes its content width count as 0, which would collapse
   the fit-content / shrink-wrap sizing above. */
#mainEdit .card-body .card:not(:has(#col3Info)) { container-type: inline-size; }

/* Card headers: floats -> flex so buttons wrap instead of overlapping */
#mainEdit .card-header:has(> .editTitle) {
    display: flex !important;
    flex-wrap: wrap !important;
    align-items: center !important;
    gap: 6px 10px !important;

    & > #pm-copy-row-btn { order: 0 !important; }
    & > .editTitle {
        order: 1 !important;
        float: none !important;
        margin: 0 auto 0 0 !important;
    }
    & > div,
    & > .float-end {
        order: 2 !important;
        float: none !important;
    }
    & .float-end {
        float: none !important;
        display: flex !important;
        flex-wrap: wrap !important;
        align-items: center !important;
        justify-content: flex-end !important;
        gap: 6px !important;
    }
    & .float-end .btn { margin: 0 !important; }
}

/* ---- Client Information table: 4-up grid, collapses to 2-up ---- */
#mainEdit .card-body > table.table-sm {
    margin: 0 !important;

    & > tbody { display: block !important; }
    & tr {
        display: grid !important;
        grid-template-columns:
            minmax(0, 0.8fr) minmax(0, 1.2fr) minmax(0, 0.8fr) minmax(0, 1.2fr) !important;
    }
    & td {
        width: auto !important;
        overflow-wrap: anywhere !important;
    }
    & td[colspan] { grid-column: span 3 !important; }
}
@container (max-width: 520px) {
    #mainEdit .card-body > table.table-sm tr {
        grid-template-columns: minmax(0, 0.9fr) minmax(0, 1.1fr) !important;
    }
    #mainEdit .card-body > table.table-sm td[colspan] { grid-column: auto !important; }
}

/* ---- Notes thread ---- */
#userNotes .notesTable {
    display: grid !important;
    grid-template-columns: 140px minmax(0, 1fr) !important;
    margin: 8px 0 0 !important;

    & > * {
        width: auto !important;
        max-width: none !important;
        margin: 0 !important;
    }
    & > .col-sm-2 {
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        padding: 10px 8px !important;
    }
    & .content .row {
        display: flex !important;
        flex-wrap: wrap !important;
        align-items: flex-start !important;
        gap: 8px !important;
        margin: 0 !important;
    }
    /* note text */
    & .content .row > .col-sm-8 {
        flex: 1 1 280px !important;
        width: auto !important;
        max-width: none !important;
        padding: 8px 0 !important;
        overflow-wrap: anywhere !important;
    }
    /* author/date box - Bootstrap's w-25 squeezed this to nothing */
    & .content .row > .col-sm-4 {
        flex: 0 1 230px !important;
        width: auto !important;
        max-width: 100% !important;
        margin: 6px 0 !important;
        padding: 6px 8px !important;
    }
}
@container (max-width: 640px) {
    #userNotes .notesTable { grid-template-columns: minmax(0, 1fr) !important; }
    #userNotes .notesTable > .col-sm-2 {
        justify-content: flex-start !important;
        padding: 6px 8px !important;
    }
}

/* ---- Save / Close footer: pinned to the bottom of the screen ---- */
#editFooter {
    position: sticky !important;
    bottom: 0 !important;
    z-index: 20 !important;
    background: var(--oes-bg) !important;
    border-top: 1px solid var(--oes-line) !important;
    padding: 8px 12px !important;

    & .buttonwrapper {
        display: flex !important;
        flex-wrap: wrap !important;
        align-items: center !important;
        gap: 8px 16px !important;
    }
    & #HighPriority,
    & #saving {
        float: none !important;
        position: static !important;
        margin: 0 !important;
    }
    & #saveCancelButtons:not([style*="none"]) {
        margin-left: auto !important;
        display: flex !important;
        flex-direction: row-reverse !important;   /* same visual order the floats gave */
        gap: 8px !important;
    }
    & .footerButton {
        float: none !important;
        margin: 0 !important;
    }
}

/* ============================================================
   5. ExtJS windows (OES Mngmt, Add Note, Document Routing, ...)
      ExtJS stamps px width/height/left/top on the window AND on every
      inner wrapper. We keep the window's own width as a preference,
      cap it to the viewport, let height follow content, and turn the
      internals into a normal flex column.
   ============================================================ */
/* ExtJS draws each window's drop shadow as a SEPARATE element, sized
   and positioned once when the window opens. When we let the window
   grow to fit its content or nudge it back on-screen, that shadow
   stays behind at the old size/spot until a drag re-syncs it - the
   gray slab sticking out past the window. Kill it and put a real
   shadow on the window itself, which always follows. */
.x-css-shadow,
.x-shadow {
    display: none !important;
}

.x-window:not(.x-hidden-offsets):not(.x-window-ghost):not([style*="display: none"]):not([style*="display:none"]) {
    box-shadow: 0 6px 24px rgba(0, 0, 0, 0.28) !important;
    display: flex !important;
    flex-direction: column !important;
    max-width: calc(100vw - 2 * var(--oes-edge)) !important;
    max-height: calc(100vh - 2 * var(--oes-edge)) !important;

    &:not(.x-resizable) { height: auto !important; }

    & > .x-window-header {
        position: relative !important;
        left: auto !important;
        top: auto !important;
        width: auto !important;
        flex: 0 0 auto !important;
    }
    & > .x-window-header .x-box-inner {
        width: auto !important;
        height: auto !important;
        overflow: visible !important;
    }
    & > .x-window-header .x-box-target {
        position: relative !important;
        display: flex !important;
        align-items: center !important;
        gap: 6px !important;
        width: auto !important;
    }
    & > .x-window-header .x-box-item {
        position: relative !important;
        left: auto !important;
        top: auto !important;
        margin: 0 !important;
    }
    & > .x-window-header .x-title {
        flex: 1 1 auto !important;
        width: auto !important;
        min-width: 0 !important;
    }
    & > .x-window-header .x-tool { flex: 0 0 auto !important; }

    & > .x-window-bodyWrap {
        display: flex !important;
        flex-direction: column !important;
        flex: 1 1 auto !important;
        min-height: 0 !important;
    }
    & .x-window-body {
        position: relative !important;
        left: auto !important;
        top: auto !important;
        width: auto !important;
        height: auto !important;
        flex: 1 1 auto !important;
        min-height: 0 !important;
        overflow: auto !important;
        container: oeswin / inline-size;
    }
    & .x-window-body > .x-panel,
    & .x-window-body .x-panel-bodyWrap,
    & .x-window-body .x-panel-body {
        position: relative !important;
        left: auto !important;
        top: auto !important;
        width: auto !important;
        height: auto !important;
    }
    & .x-window-body .x-panel-body { overflow: visible !important; }
    & .x-window-body .x-autocontainer-outerCt {
        width: 100% !important;
        height: auto !important;
    }
    & select,
    & input.oesInput,
    & textarea { max-width: 100% !important; }
}

/* ---- 5a. OES Management window ---- */
#OESMngmtWin-body {
    height: auto !important;

    & .tab-pane { overflow-x: auto !important; }
    & #mngmtTab1.active {
        display: grid !important;
        grid-template-columns: minmax(0, 1.7fr) minmax(220px, 1fr) !important;
        align-items: start !important;
        gap: var(--oes-gap) !important;
        padding-top: 8px !important;
    }
    & #mngmtMain,
    & #mngmtDocs {
        position: static !important;
        float: none !important;
        width: auto !important;
        height: auto !important;
        inset: auto !important;
        margin: 0 !important;
        min-width: 0 !important;
    }
    & #mngmtMain > table,
    & #mngmtMain > table > tbody {
        display: block !important;
        width: 100% !important;
    }
    & #mngmtMain tr:not([style*="none"]) {
        display: flex !important;
        flex-wrap: wrap !important;
        gap: 6px 12px !important;
        margin-bottom: 6px !important;
    }
    & #mngmtMain td {
        display: block !important;
        flex: 1 1 220px !important;
        min-width: 0 !important;
        padding: 0 !important;
    }
    & #mngmtMain .oesInput {
        width: 100% !important;
        box-sizing: border-box !important;
    }
    & #cliName:not([style*="none"]) {
        display: flex !important;
        gap: 6px !important;
    }
    & #cliName .oesInput { width: auto !important; min-width: 0 !important; }
    & #cliName [name="clientFirst"]  { flex: 3 1 0 !important; }
    & #cliName [name="clientMiddle"] { flex: 2 1 0 !important; }
    & #cliName [name="clientLast"]   { flex: 4 1 0 !important; }

    & #MngmtMainBttm {
        grid-column: 1 / -1 !important;
        position: static !important;
        display: flex !important;
        align-items: center !important;
        justify-content: space-between !important;
        gap: 8px !important;
        width: auto !important;
        padding-top: 8px !important;
        border-top: 1px solid var(--oes-line) !important;
    }
    & #MngmtMainBttm > *,
    & #mngmtMainSaveBtn {
        float: none !important;
        margin: 0 !important;
    }
    & .mngmtSaving {
        grid-column: 1 / -1 !important;
        justify-self: end !important;
    }
    & #mngmtTab4 > div[style*="350px"] {
        height: auto !important;
        max-height: 55vh !important;
    }
}
@container oeswin (max-width: 620px) {
    #OESMngmtWin-body #mngmtTab1.active { grid-template-columns: minmax(0, 1fr) !important; }
}

/* ---- 5b. Add Note / NIGO note window ---- */
#noteWin-body {
    height: auto !important;

    & > div { height: auto !important; }
    & form[name="noteForm"] {
        display: flex !important;
        flex-direction: column !important;
        gap: 8px !important;
        padding: 6px !important;
    }
    & form[name="noteForm"] > div {
        margin: 0 !important;
        padding: 0 !important;
    }
    & textarea.addOESNotes {
        width: 100% !important;
        height: auto !important;
        min-height: 150px !important;
        resize: vertical !important;
        box-sizing: border-box !important;
    }
    & select[name="docID"] {
        width: 100% !important;
        margin: 0 !important;
    }
    /* the Save/Cancel box was absolutely pinned and overlapped the quick-fill buttons */
    & form[name="noteForm"] > div[style*="position:absolute"] {
        position: static !important;
        align-self: flex-end !important;
    }
    & .buttonwrapper {
        display: flex !important;
        flex-direction: row-reverse !important;
        gap: 8px !important;
    }
    & .buttonwrapper .btn {
        float: none !important;
        margin: 0 !important;
    }
}

/* ---- 5c. Document Routing window ---- */
.routingWin {
    height: auto !important;

    & form[name="docRoutingForm"] {
        display: grid !important;
        grid-template-columns: minmax(0, 1fr) minmax(0, 1fr) !important;
        gap: 10px !important;
        padding: 8px 32px 8px 8px !important;   /* right padding clears the close X */
    }
    /* "Route Documents to:" row */
    & form > div[style*="position:absolute"] {
        grid-column: 1 / -1 !important;
        position: static !important;
        display: flex !important;
        flex-wrap: wrap !important;
        align-items: center !important;
        gap: 8px !important;
    }
    & select.routeDocsTo {
        flex: 1 1 240px !important;
        width: auto !important;
    }
    & #routingDocs,
    & #routingNotes {
        position: static !important;
        float: none !important;
        inset: auto !important;
        width: auto !important;
        height: auto !important;
        margin: 0 !important;
        min-width: 0 !important;
    }
    & #routingNotes {
        display: flex !important;
        flex-direction: column !important;
        gap: 6px !important;
    }
    & #routingNotes textarea {
        width: 100% !important;
        height: auto !important;
        min-height: 180px !important;
        resize: vertical !important;
    }
    & #saveRoutingBtn,
    & #routeSaving {
        position: static !important;
        align-self: flex-end !important;
    }
}
@container oeswin (max-width: 560px) {
    .routingWin form[name="docRoutingForm"] { grid-template-columns: minmax(0, 1fr) !important; }
}

/* ============================================================
   6. Other popup systems
   ============================================================ */
.ui-dialog {
    max-width: calc(100vw - 2 * var(--oes-edge)) !important;
}
.ui-dialog .ui-dialog-content {
    max-height: calc(100vh - 120px) !important;
    overflow: auto !important;
}
.jBox-wrapper .jBox-container {
    max-width: calc(100vw - 2 * var(--oes-edge)) !important;
    max-height: calc(100vh - 2 * var(--oes-edge)) !important;
    overflow: auto !important;
}

/* ============================================================
   7. Click-to-copy OES IDs
      Plain-text IDs get a pointer cursor; on click the text
      briefly flashes gray as confirmation.
   ============================================================ */
#editHeader .row:first-child > :first-child,
#mainEdit .docTable td.title,
#OESMngmtWin-body #mngmtTab4 > div:first-child,
#lockedOES {
    cursor: pointer !important;
}

@keyframes oesCopyFlash {
    from { color: #9a9a9a; }
}
.oes-copy-flash,
.oes-copy-flash * {
    animation: oesCopyFlash 450ms ease-out;
}
@media (prefers-reduced-motion: reduce) {
    .oes-copy-flash,
    .oes-copy-flash * { animation-duration: 1ms; }
}

} /* end @layer oesfix */
`;

    // ---------------------------------------------------------------
    // Inject stylesheet as early as possible (no flash of broken layout)
    // ---------------------------------------------------------------
    const style = document.createElement('style');
    style.id = 'oes-responsive-layout';
    style.textContent = CSS;
    (document.head || document.documentElement).appendChild(style);

    // ---------------------------------------------------------------
    // Keep popups on screen.
    // ExtJS / jQuery UI position windows once, in px, based on the size
    // the window was at that moment. We nudge them back inside the
    // viewport whenever they move, resize, or the browser resizes.
    // Dragging still works; you just can't lose a window off-screen.
    // ---------------------------------------------------------------
    const EDGE = 8;
    const POPUP_SEL = '.x-window:not(.x-window-ghost), .ui-dialog, .jBox-wrapper';

    function isShown(el) {
        if (el.classList.contains('x-hidden-offsets')) return false;
        const cs = getComputedStyle(el);
        return cs.display !== 'none' && cs.visibility !== 'hidden';
    }

    function clamp(el) {
        if (style.disabled || !el.isConnected || !isShown(el)) return;
        const r = el.getBoundingClientRect();
        if (!r.width || !r.height) return;

        const maxLeft = Math.max(EDGE, window.innerWidth - r.width - EDGE);
        const maxTop = Math.max(EDGE, window.innerHeight - r.height - EDGE);
        const dx = Math.min(Math.max(r.left, EDGE), maxLeft) - r.left;
        const dy = Math.min(Math.max(r.top, EDGE), maxTop) - r.top;
        if (Math.abs(dx) < 1 && Math.abs(dy) < 1) return;

        // Delta-based, so it works whether the element is fixed or
        // absolute and whatever its offset parent is.
        const cs = getComputedStyle(el);
        el.style.left = (parseFloat(cs.left) || 0) + dx + 'px';
        el.style.top = (parseFloat(cs.top) || 0) + dy + 'px';
    }

    const pending = new Set();
    let raf = 0;
    function schedule(el) {
        pending.add(el);
        if (raf) return;
        raf = requestAnimationFrame(() => {
            raf = 0;
            pending.forEach(clamp);
            pending.clear();
        });
    }

    const watched = new WeakSet();
    const attrObserver = new MutationObserver((ms) => ms.forEach((m) => schedule(m.target)));
    const sizeObserver = new ResizeObserver((entries) => entries.forEach((e) => schedule(e.target)));

    function watch(el) {
        if (watched.has(el)) return;
        watched.add(el);
        attrObserver.observe(el, { attributes: true, attributeFilter: ['style', 'class'] });
        sizeObserver.observe(el);
        schedule(el);
    }

    function scan(root) {
        if (root.matches && root.matches(POPUP_SEL)) watch(root);
        if (root.querySelectorAll) root.querySelectorAll(POPUP_SEL).forEach(watch);
    }

    function rescanAll() {
        document.querySelectorAll(POPUP_SEL).forEach(schedule);
    }

    function init() {
        scan(document.body);
        // All three popup systems render straight into <body>, so we only
        // need direct children, not the whole (very chatty) subtree.
        new MutationObserver((ms) => {
            for (const m of ms) {
                for (const n of m.addedNodes) {
                    if (n.nodeType === 1) scan(n);
                }
            }
        }).observe(document.body, { childList: true });
        window.addEventListener('resize', rescanAll);
    }

    // ---------------------------------------------------------------
    // Click-to-copy OES IDs.
    // One capture-phase listener on the document, so it covers content
    // the site loads later via AJAX without re-scanning anything. It
    // never cancels the click.
    //
    // Copies:
    //  - open-OES header "OES ID: 1234567"
    //  - OES Mngmt Activity tab "OES ID: 1234567"
    //  - window / dialog titles "OES #1234567", "Scans for OES #1234567"
    //  - Related Client Profiles OES number in the Documents card
    //  - the lock badge (#lockedOES)
    //
    // Does NOT copy:
    //  - links that open an OES (queue rows, Recently Viewed, OES Mngmt
    //    "Other OES") - clicking those just navigates, as normal
    //  - R/O ID links
    // ---------------------------------------------------------------
    const ID_RE = /^\d{5,9}$/;
    const LABELED_RE = /OES\s*(?:ID)?\s*[:#]\s*(\d{5,9})/i;

    function cleanText(el) {
        return (el.textContent || '').replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();
    }

    // Returns { id, el } where el is the element to flash, or null.
    function findOesId(target) {
        if (!(target instanceof Element)) return null;

        // Any link that does something (opens an OES, opens a rollover,
        // etc.) is left alone.
        if (target.closest('a[onclick], a[href]')) return null;

        // Lock badge
        const lock = target.closest('#lockedOES');
        if (lock) {
            const t = cleanText(lock);
            return ID_RE.test(t) ? { id: t, el: lock } : null;
        }

        // Related Client Profiles header cell in the Documents card
        const td = target.closest('.docTable td.title');
        if (td) {
            const t = cleanText(td);
            return ID_RE.test(t) ? { id: t, el: td } : null;
        }

        // Labeled text: "OES ID: 1234567", "OES #1234567".
        // Walk up a few short ancestors only, so clicking somewhere big
        // that merely *contains* an OES ID doesn't copy.
        for (let el = target, hops = 0; el && el !== document.body && hops < 4; el = el.parentElement, hops++) {
            const t = cleanText(el);
            if (t.length > 60) break;
            const m = t.match(LABELED_RE);
            if (m) return { id: m[1], el };
        }
        return null;
    }

    function flash(el) {
        el.classList.remove('oes-copy-flash');
        void el.offsetWidth;                    // restart the animation on repeat clicks
        el.classList.add('oes-copy-flash');
        clearTimeout(el._oesFlashTimer);
        el._oesFlashTimer = setTimeout(() => el.classList.remove('oes-copy-flash'), 500);
    }

    function copyText(text) {
        if (typeof GM_setClipboard === 'function') {
            GM_setClipboard(text, { type: 'text', mimetype: 'text/plain' });
        } else if (navigator.clipboard) {
            navigator.clipboard.writeText(text);
        }
    }

    document.addEventListener('click', (e) => {
        if (e.button !== 0) return;
        // Don't stomp on a text selection you're making by hand
        const sel = window.getSelection && window.getSelection().toString();
        if (sel) return;

        const hit = findOesId(e.target);
        if (!hit) return;
        copyText(hit.id);
        flash(hit.el);
    }, true);

    // ---------------------------------------------------------------
    // Reopen where you were after a page reload.
    //
    // No simulated clicks. Every OES link runs showEdit(...), every tab
    // runs changeState(n). We wrap those two functions (plus closeEdit)
    // so each call is remembered in sessionStorage, then on a real
    // reload we call the same functions again with the same arguments.
    //
    //  - sessionStorage is per-tab: survives reload, gone when the tab
    //    closes. A new tab / fresh login never auto-opens anything.
    //  - Only fires on an actual reload (F5 / Ctrl+R), not on normal
    //    navigation, and only if the saved state is < 30 min old.
    //  - If the module differs, it calls changeState() first, then
    //    reopens the OES once that page has loaded.
    //  - If the OES doesn't open within 8s it gives up quietly and
    //    you're left wherever the site put you, same as today.
    //
    // Runs in the page's own JS context (injected <script>) because it
    // has to wrap and call the site's functions.
    // Set to false to turn it off.
    // ---------------------------------------------------------------
    const RESTORE_ON_RELOAD = true;

    function pageRestore() {
        var KEY = 'oesfix:last';
        var PENDING = 'oesfix:pending';
        var MAX_AGE = 30 * 60 * 1000;

        function read(k) {
            try { return JSON.parse(sessionStorage.getItem(k)); } catch (e) { return null; }
        }
        function write(k, v) {
            try {
                if (v == null) sessionStorage.removeItem(k);
                else sessionStorage.setItem(k, JSON.stringify(v));
            } catch (e) {}
        }
        function log() {
            try { console.info.apply(console, ['[OES restore]'].concat([].slice.call(arguments))); } catch (e) {}
        }
        function currentModule() {
            var el = document.getElementById('openModule');
            return el ? String(el.value) : null;
        }
        function remember(oesArgs) {
            write(KEY, { module: currentModule(), oes: oesArgs || null, ts: Date.now() });
        }
        function navType() {
            var n = performance.getEntriesByType && performance.getEntriesByType('navigation')[0];
            return n ? n.type : 'navigate';
        }

        function wrap(name, before) {
            var orig = window[name];
            if (typeof orig !== 'function' || orig.__oesfix) return;
            var wrapped = function () {
                try { before.apply(this, arguments); } catch (e) {}
                return orig.apply(this, arguments);
            };
            wrapped.__oesfix = true;
            window[name] = wrapped;
        }
        function wrapAll() {
            wrap('showEdit', function () { remember([].slice.call(arguments)); });
            wrap('closeEdit', function () { remember(null); });
            wrap('changeState', function (n) {
                write(KEY, { module: String(n), oes: null, ts: Date.now() });
            });
        }

        // Wait for the page (and the site's own startup code) to settle
        function whenReady(fn) {
            function go() { setTimeout(fn, 400); }
            if (document.readyState === 'complete') go();
            else window.addEventListener('load', go, { once: true });
        }

        function reopen(args) {
            var tries = 0;
            (function attempt() {
                if (typeof window.showEdit !== 'function') {
                    if (++tries < 20) setTimeout(attempt, 250);
                    return;
                }
                log('reopening OES', args[0]);
                window.showEdit.apply(window, args);

                var started = Date.now();
                (function verify() {
                    var idEl = document.getElementById('OESID');
                    if (idEl && String(idEl.value) === String(args[0])) return;   // opened
                    if (Date.now() - started > 8000) {
                        log('could not reopen OES', args[0], '- giving up');
                        remember(null);
                        return;
                    }
                    setTimeout(verify, 250);
                })();
            })();
        }

        function start() {
            wrapAll();

            var mod = currentModule();
            var saved = read(KEY);
            var pending = read(PENDING);

            // Phase 2: we already switched modules on the previous load
            if (pending) {
                write(PENDING, null);
                var st = pending.state;
                if (st && st.oes && mod === st.module) {
                    write(KEY, st);
                    whenReady(function () { reopen(st.oes); });
                } else {
                    remember(null);
                }
                return;
            }

            var fresh = saved && (Date.now() - saved.ts) < MAX_AGE;

            if (navType() === 'reload' && fresh) {
                // Phase 1a: wrong module - go back to the right one first
                if (saved.module && mod !== saved.module && typeof window.changeState === 'function') {
                    write(PENDING, { state: saved });
                    log('returning to module', saved.module);
                    whenReady(function () {
                        window.changeState(Number(saved.module));
                        // In case changeState swaps content without a full
                        // page load, pick it up here instead of on next load
                        if (saved.oes) {
                            setTimeout(function () {
                                if (read(PENDING) && currentModule() === saved.module) {
                                    write(PENDING, null);
                                    write(KEY, saved);
                                    reopen(saved.oes);
                                }
                            }, 3000);
                        }
                    });
                    return;
                }
                // Phase 1b: right module, just reopen the OES
                if (saved.oes && mod === saved.module) {
                    whenReady(function () { reopen(saved.oes); });
                    return;
                }
            }

            // Ordinary page load: remember where we are, nothing open yet
            remember(null);
        }

        document.addEventListener('DOMContentLoaded', start);
        // Belt and braces in case the site defines these late
        window.addEventListener('load', wrapAll);
    }

    if (RESTORE_ON_RELOAD) {
        const s = document.createElement('script');
        s.textContent = '(' + pageRestore.toString() + ')();';
        (document.head || document.documentElement).appendChild(s);
        s.remove();
    }

    // Alt+Shift+R: toggle the layout fix to compare against the original
    document.addEventListener('keydown', (e) => {
        if (e.altKey && e.shiftKey && e.code === 'KeyR') {
            style.disabled = !style.disabled;
            console.info('[OES Responsive] ' + (style.disabled ? 'OFF' : 'ON'));
            if (!style.disabled) rescanAll();
        }
    }, true);

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();