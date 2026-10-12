// ==UserScript==
// @name         OES Trade Ticket — ACH Checker (Purchase Team)
// @namespace    http://tampermonkey.net/
// @version      2.0
// @description  Opens Trade Ticket PDFs in a shell window with a validation panel and a formatted copy-to-clipboard string
// @match        https://oes.example.com/
// @match        https://oes.example.com/index.cfm*
// @require      https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js
// @grant        none
// @run-at       document-idle
// ==/UserScript==

(function () {
    'use strict';

    // ===== CONFIG ========================================================
    const NAV_STATE   = 56;                 // Purchase Team -> changeState(56)
    const DOC_NAMES   = ['trade ticket'];   // lowercase
    const MARKER      = 'q=runuser';
    const WORKER_URL  = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

    const CATCH_UP_AGE = 50;                // age >= this uses the catch-up limit
    const LIMITS = {                        // by tax year; add a line each January
        2026: { base: 7500, catchUp: 8600 },
    };
    const DISCLOSURE_LIMIT = 250000;   // at or above this, disclosure must be confirmed
    // Account types the contribution rule applies to (lowercase, substring match)
    const LIMITED_TYPES = ['ira', 'roth ira'];

    // Output: OES # 1234567 Lastname, First RN 123456789 AN 987654321 Checking $50.00
    const NAME_AS_IS = false;  // true -> leave the PDF's "First Last" order alone
    const KEEP_SUFFIX = false; // true -> "Brown, William Jr" instead of "Brown, William"

    const DEBUG = false;
    // =====================================================================

    const log = (...a) => DEBUG && console.log('[TT]', ...a);

    // ---------- nav / row guards ----------------------------------------
    function navIsActive() {
        const active = document.querySelector('#nav .navBox.activeNav');
        if (!active) return false;
        const re = new RegExp('changeState\\(\\s*' + NAV_STATE + '\\s*\\)');
        return re.test(active.getAttribute('onclick') || '');
    }

    function rowDocName(tr) {
        const a = tr.querySelector('td:first-child a.clickable');
        return a ? a.textContent.trim().toLowerCase() : '';
    }

    function docIDFrom(tr) {
        const m = tr.innerHTML.match(/docWindow\((\d+)\)/);
        return m ? m[1] : null;
    }

    function getOESNumber() {
        const re = /OES\s*(?:ID|#)\s*:?\s*(\d+)/i;
        const header = document.querySelector('#editHeader');
        let m = header && (header.innerText || '').match(re);
        if (m) return m[1];
        m = (document.body.innerText || '').match(re);
        return m ? m[1] : '';
    }

    // ---------- click interception --------------------------------------
    document.addEventListener('click', (e) => {
        if (!e.target.closest) return;
        const trigger = e.target.closest('a.clickable, img.docLink');
        if (!trigger) return;

        const tr = trigger.closest('tr');
        if (!tr || !tr.closest('table.docTable')) return;
        if (!DOC_NAMES.includes(rowDocName(tr))) return;
        if (!navIsActive()) return;

        const docID = docIDFrom(tr);
        if (!docID) return;

        // take over from docWindow()
        e.preventDefault();
        e.stopImmediatePropagation();

        openChecker(docID, getOESNumber());
    }, true);

    // ---------- shell window --------------------------------------------
    const SHELL_CSS = `
        * { box-sizing: border-box; }
        body { margin:0; font:13px/1.45 -apple-system,Segoe UI,Roboto,sans-serif;
               display:flex; height:100vh; background:#f4f4f5; color:#18181b; }
        #pdf { flex:1 1 auto; border:0; background:#525659; }
        #panel { flex:0 0 380px; overflow:auto; padding:16px; background:#fff;
                 border-left:1px solid #d4d4d8; }
        h2 { margin:0 0 4px; font-size:15px; }
        .sub { color:#71717a; font-size:11px; margin-bottom:14px; }
        .field { display:flex; align-items:center; gap:8px; margin-bottom:6px; }
        .field label { flex:0 0 96px; color:#71717a; font-size:11px;
                       text-transform:uppercase; letter-spacing:.04em; }
        .field input { flex:1 1 auto; padding:5px 7px; border:1px solid #d4d4d8;
                       border-radius:4px; font:12px ui-monospace,Menlo,monospace; }
        .field input:read-only { background:#fafafa; color:#3f3f46; }
        hr { border:0; border-top:1px solid #e4e4e7; margin:14px 0; }
        .check { display:flex; gap:8px; align-items:flex-start; margin-bottom:8px; }
        .mark { flex:0 0 16px; font-weight:700; line-height:1.4; }
        .pass .mark { color:#16a34a; }
        .fail .mark { color:#dc2626; }
        .skip .mark { color:#a1a1aa; }
        .check .body { flex:1 1 auto; }
        .check .why { color:#71717a; font-size:11px; }
        textarea { width:100%; height:76px; padding:8px; margin-top:4px;
                   border:1px solid #d4d4d8; border-radius:4px; resize:vertical;
                   font:12px ui-monospace,Menlo,monospace; }
        button { width:100%; padding:9px; margin-top:8px; border:0; border-radius:5px;
                 background:#18181b; color:#fff; font-size:13px; cursor:pointer; }
        button:hover { background:#3f3f46; }
        button.ghost { background:#e4e4e7; color:#3f3f46; margin-top:6px; }
        #status { padding:8px 10px; border-radius:4px; background:#fef9c3; font-size:12px; }
        #dump { display:none; white-space:pre; font:10px ui-monospace,monospace;
                background:#fafafa; border:1px solid #e4e4e7; border-radius:4px;
                padding:8px; margin-top:8px; max-height:320px; overflow:auto; }
        #panelHead { display:flex; justify-content:space-between; align-items:flex-start; }
        #hideBtn { width:auto; margin:0; padding:2px 8px; background:#e4e4e7; color:#3f3f46;
                   font-size:15px; line-height:1.2; border-radius:4px; }
        #notesWrap { margin-top:10px; }
        #notesWrap .cap { font-size:11px; color:#71717a; text-transform:uppercase;
                          letter-spacing:.04em; margin-bottom:4px; }
        #notes { background:#000; color:#facc15; padding:9px 10px; border-radius:4px;
                 font:12px/1.5 ui-monospace,Menlo,monospace; white-space:pre-wrap;
                 border:2px solid #facc15; }
        #notes.empty { color:#71717a; border-color:#3f3f46; font-style:italic; }
        #bigWarn { display:none; margin-top:12px; padding:12px; border-radius:5px;
                   background:#fef2f2; border:2px solid #dc2626; }
        #bigWarn h3 { margin:0 0 6px; font-size:14px; color:#b91c1c; }
        #bigWarn p { margin:0 0 8px; font-size:12px; color:#7f1d1d; }
        #bigWarn label { display:flex; gap:7px; align-items:flex-start;
                         font-size:12px; color:#7f1d1d; cursor:pointer; }
        #blockMsg { display:none; margin-top:8px; padding:9px 10px; border-radius:4px;
                    background:#fef2f2; border:1px solid #dc2626; color:#b91c1c; font-size:12px; }
        button:disabled { background:#d4d4d8 !important; color:#71717a !important;
                          cursor:not-allowed; }
        #reopen { display:none; position:fixed; right:0; top:50%; transform:translateY(-50%);
                  background:#18181b; color:#fff; padding:14px 7px; cursor:pointer;
                  border-radius:5px 0 0 5px; font-size:14px; user-select:none;
                  writing-mode:vertical-rl; letter-spacing:.08em; }
    `;

    const SHELL_HTML = (pdfURL) => `
        <iframe id="pdf" src="${pdfURL}"></iframe>
        <div id="panel">
            <div id="panelHead">
                <h2>Trade Ticket check</h2>
                <button id="hideBtn" title="Hide panel">›</button>
            </div>
            <div class="sub" id="sub">reading PDF…</div>
            <div id="status">Extracting…</div>
            <div id="fields" style="display:none">
                <div class="field"><label>OES #</label><input id="f_oes"></div>
                <div class="field"><label>Name</label><input id="f_name"></div>
                <div class="field"><label>RN</label><input id="f_rn"></div>
                <div class="field"><label>AN</label><input id="f_an"></div>
                <div class="field"><label>Bank type</label><input id="f_type"></div>
                <div class="field"><label>Amount</label><input id="f_amt"></div>
                <hr>
                <div id="checks"></div>
                <hr>
                <textarea id="out" readonly></textarea>
                <div id="bigWarn">
                    <h3>Large amount — $250,000 or more</h3>
                    <p>Confirm this has been disclosed on a subsequent form before copying.</p>
                    <label><input type="checkbox" id="ackBox">
                        <span>I've confirmed the disclosure</span></label>
                </div>
                <div id="blockMsg"></div>
                <button id="copyBtn">Copy to clipboard</button>
                <button id="dumpBtn" class="ghost">Show extracted text</button>
                <pre id="dump"></pre>
                <div id="notesWrap">
                    <div class="cap">Section Five notes — read these</div>
                    <div id="notes"></div>
                </div>
            </div>
        </div>
        <div id="reopen" title="Show panel">CHECKS</div>
    `;

    function openChecker(docID, oes) {
        const pdfURL = `${location.origin}/common/loadDoc.cfm?docID=${docID}` +
                       `&showNav=0&page=1&jobID=0&${MARKER}`;

        const win = window.open('', 'ttcheck_' + docID,
                                'width=1500,height=950,scrollbars=yes,resizable=yes');
        if (!win) { alert('Popup blocked — allow popups for this site.'); return; }

        win.document.open();
        win.document.write(
            `<!DOCTYPE html><html><head><meta charset="utf-8">` +
            `<title>Trade Ticket ${docID}</title><style>${SHELL_CSS}</style></head>` +
            `<body>${SHELL_HTML(pdfURL)}</body></html>`
        );
        win.document.close();

        // panel hide / show
        const panel  = win.document.getElementById('panel');
        const reopen = win.document.getElementById('reopen');
        const hide   = win.document.getElementById('hideBtn');
        hide.addEventListener('click', () => {
            panel.style.display = 'none';
            reopen.style.display = 'block';
        });
        reopen.addEventListener('click', () => {
            panel.style.display = 'block';
            reopen.style.display = 'none';
        });

        run(win, pdfURL, docID, oes).catch(err => {
            console.error('[TT]', err);
            const s = win.document.getElementById('status');
            if (s) s.textContent = 'Failed: ' + err.message;
        });
    }

    // ---------- pdf.js ---------------------------------------------------
    let workerReady = null;
    async function initWorker() {
        if (workerReady) return workerReady;
        workerReady = (async () => {
            try {
                const res = await fetch(WORKER_URL);
                const code = await res.text();
                const blob = new Blob([code], { type: 'application/javascript' });
                pdfjsLib.GlobalWorkerOptions.workerSrc = URL.createObjectURL(blob);
            } catch (e) {
                log('blob worker failed, falling back to CDN url', e);
                pdfjsLib.GlobalWorkerOptions.workerSrc = WORKER_URL;
            }
        })();
        return workerReady;
    }

    async function extractItems(buf) {
        await initWorker();
        const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
        const items = [];
        for (let p = 1; p <= pdf.numPages; p++) {
            const page = await pdf.getPage(p);
            const vp = page.getViewport({ scale: 1 });
            const tc = await page.getTextContent();
            for (const it of tc.items) {
                const str = (it.str || '').replace(/\s+/g, ' ').trim();
                if (!str) continue;
                items.push({
                    str,
                    page: p,
                    x: it.transform[4],
                    y: vp.height - it.transform[5],   // y from top of page
                    w: it.width || 0
                });
            }
        }
        // reading order
        items.sort((a, b) => a.page - b.page || a.y - b.y || a.x - b.x);
        return items;
    }

    // ---------- layout helpers ------------------------------------------
    const ROW_TOL = 4;

    function buildRows(items) {
        const rows = [];
        for (const it of items) {
            const row = rows.find(r => r.page === it.page && Math.abs(r.y - it.y) <= ROW_TOL);
            if (row) { row.items.push(it); row.y = (row.y + it.y) / 2; }
            else rows.push({ page: it.page, y: it.y, items: [it] });
        }
        rows.forEach(r => r.items.sort((a, b) => a.x - b.x));
        rows.sort((a, b) => a.page - b.page || a.y - b.y);
        return rows;
    }

    function tagSections(items) {
        let current = '';
        for (const it of items) {
            const m = it.str.match(/^SECTION\s+([A-Z]+)/i);
            if (m) current = m[1].toUpperCase();
            it.section = current;
        }
    }

    // Anything that looks like a form label shouldn't be harvested as a value.
    const LABELS = [
        'date','call originator','caller','security review confirmed','service center rep',
        'call log id','solicitation','ssn','client name','client age','registered rep',
        'rep id','account number','redemption','account type','product','portfolio','balance',
        'cost basis','change','full account distribution','notes','partial internal transfer',
        'full internal transfer','exchange money between funds','ticker','fund #','rebalance my account',
        'amount','tax year','ach bank info on file?','bank name','routing number','employer',
        'bank account type','contra firm','contra firm contact','telephone','stop pay method',
        'check number','check date','check amount','reason for stop pay','yes','no','n/a'
    ];
    const squash = s => s.toLowerCase().replace(/\s+/g, '');
    const matchesLabel = (str, label) => squash(str) === squash(label);
    const isLabel = s => LABELS.some(l => squash(l) === squash(s)) ||
                         /^SECTION\s/i.test(s);

    /**
     * Find a value for `label`, trying in order:
     *   A. items to the right of the label on the same row
     *   B. items directly beneath the label (column layout)
     *   C. any item in the same section that passes `validate`
     */
    function getField(rows, items, label, opts = {}) {
        const { section, validate = s => !isLabel(s), claimed = new Set() } = opts;
        const sections = section ? [].concat(section) : null;
        const inSection = it => !sections || sections.includes(it.section);

        const rowIdx = rows.findIndex(r => {
            if (r.items[0] && !inSection(r.items[0])) return false;
            return r.items.some(i => matchesLabel(i.str, label));
        });

        if (rowIdx >= 0) {
            const row = rows[rowIdx];
            const li = row.items.findIndex(i => matchesLabel(i.str, label));
            const labelItem = row.items[li];

            // A — to the right
            for (let k = li + 1; k < row.items.length; k++) {
                const cand = row.items[k];
                if (isLabel(cand.str)) continue;
                if (claimed.has(cand)) continue;
                if (validate(cand.str)) { claimed.add(cand); return cand.str; }
            }

            // B — beneath, x-aligned
            for (let r = rowIdx + 1; r < Math.min(rowIdx + 4, rows.length); r++) {
                for (const cand of rows[r].items) {
                    if (Math.abs(cand.x - labelItem.x) > 40) continue;
                    if (isLabel(cand.str) || claimed.has(cand)) continue;
                    if (validate(cand.str)) { claimed.add(cand); return cand.str; }
                }
            }
        }

        // C — section-wide scan
        if (sections) {
            for (const cand of items) {
                if (!inSection(cand)) continue;
                if (isLabel(cand.str) || claimed.has(cand)) continue;
                if (validate(cand.str)) { claimed.add(cand); return cand.str; }
            }
        }
        return '';
    }

    // ---------- field extraction ----------------------------------------
    const V = {
        routing:  s => /^\d{9}$/.test(s),
        bankAcct: s => /^\d{4,17}$/.test(s),
        bankType: s => /^(checking|savings)$/i.test(s),
        money:    s => /^\$?[\d,]+\.\d{2}$/.test(s),
        age:      s => /^\d{1,3}(\.\d+)?$/.test(s) && parseFloat(s) < 120,
        year:     s => /^(19|20)\d{2}$/.test(s),
        name:     s => /^[A-Za-z][A-Za-z .,'\-]{2,}$/.test(s),
        acctType: s => /^[A-Za-z0-9 ()\/\-]{2,30}$/.test(s) && !/^(yes|no)$/i.test(s),
    };

    function parseFields(items) {
        tagSections(items);
        const rows = buildRows(items);
        const claimed = new Set();

        const f = {
            name:     getField(rows, items, 'Client Name',      { section:'ONE',   validate:V.name,     claimed }),
            age:      getField(rows, items, 'Client Age',       { section:['ONE','TWO'], validate:V.age, claimed }),
            acctType: getField(rows, items, 'Account Type',     { section:'FIVE',  validate:V.acctType, claimed }),
            amount:   getField(rows, items, 'Amount',           { section:'FIVE',  validate:V.money,    claimed }),
            taxYear:  getField(rows, items, 'Tax Year',         { section:'FIVE',  validate:V.year,     claimed }),
            bankName: getField(rows, items, 'Bank Name',        { section:'SEVEN', validate:s => !isLabel(s), claimed }),
            routing:  getField(rows, items, 'Routing Number',   { section:'SEVEN', validate:V.routing,  claimed }),
            bankAcct: getField(rows, items, 'Account Number',   { section:'SEVEN', validate:V.bankAcct, claimed }),
            bankType: getField(rows, items, 'Bank Account Type',{ section:'SEVEN', validate:V.bankType, claimed }),
        };
        return { fields: f, rows };
    }

    const SUFFIX_RE = /^(jr|sr|ii|iii|iv|v|md|phd|esq)\.?$/i;

    // Lowercase name particles that are genuinely separate words.
    const PARTICLES = new Set(['van','von','de','del','della','di','da','das','dos',
                               'du','den','der','ten','ter','la','le','bin','ibn','al']);

    // The PDF splits words mid-token ("W illiam Brow n"). A fragment starting with a
    // lowercase letter continues the previous token; real names start with capitals.
    function dekern(s) {
        const out = [];
        for (const tok of s.trim().split(/\s+/)) {
            const bare = tok.toLowerCase().replace(/[.,]/g, '');
            const c = tok[0] || '';
            const startsLower = c === c.toLowerCase() && c !== c.toUpperCase();
            if (out.length && startsLower && !PARTICLES.has(bare)) out[out.length - 1] += tok;
            else out.push(tok);
        }
        return out.join(' ');
    }

    // "Notes" appears in almost every section. This reads only the one in
    // SECTION FIVE — between its Notes label and the "If Representative
    // Initiated" line that closes the section.
    function sectionFiveNotes(rows) {
        const idx = rows.findIndex(r =>
            r.items[0] && r.items[0].section === 'FIVE' &&
            r.items.some(i => matchesLabel(i.str, 'Notes')));
        if (idx < 0) return '';

        const out = [];
        for (let k = idx + 1; k < rows.length; k++) {
            const row = rows[k];
            const text = row.items.map(i => i.str).join(' ').trim();
            if (!text) continue;
            if (/^SECTION\s/i.test(text)) break;
            if (/^If\s+Representative\s+Initiated/i.test(text)) break;
            if (/^If\s+Ticket\s+is\s+received/i.test(text)) break;
            if (row.items[0] && row.items[0].section !== 'FIVE') break;
            out.push(text);
        }
        return out.join(' ').replace(/\s+/g, ' ').trim();
    }

    function formatName(raw) {
        if (!raw) return '';
        let s = dekern(raw.trim().replace(/\s+/g, ' ').replace(/,\s*$/, ''));
        if (NAME_AS_IS) return s;

        // Pull a trailing suffix, comma-separated or not: "Heaton, II" / "Brown Jr"
        let suffix = '';
        const m = s.match(/[,\s]+([A-Za-z]{1,4}\.?)$/);
        if (m && SUFFIX_RE.test(m[1])) {
            suffix = m[1];
            s = s.slice(0, m.index).trim().replace(/,$/, '');
        }

        const parts = s.split(',').map(p => p.trim()).filter(Boolean);
        let flipped;
        if (parts.length >= 2) {
            flipped = s;                              // already "Last, First"
        } else {
            const i = s.indexOf(' ');
            flipped = i < 0 ? s : `${s.slice(i + 1)}, ${s.slice(0, i)}`;
        }
        return (KEEP_SUFFIX && suffix) ? `${flipped} ${suffix}` : flipped;
    }

    // ---------- rules ----------------------------------------------------
    // Each rule: { label, run(f) -> {state:'pass'|'fail'|'skip', why} }
    const RULES = [
        {
            label: 'All required fields found',
            run: f => {
                const missing = ['name','routing','bankAcct','bankType','amount']
                    .filter(k => !f[k]);
                return missing.length
                    ? { state:'fail', why:'missing: ' + missing.join(', ') }
                    : { state:'pass', why:'' };
            }
        },
        {
            label: 'Routing number is 9 digits',
            run: f => !f.routing ? { state:'skip', why:'not found' }
                : V.routing(f.routing) ? { state:'pass', why:'' }
                : { state:'fail', why:`got "${f.routing}"` }
        },
        {
            label: 'Bank account type valid',
            run: f => !f.bankType ? { state:'skip', why:'not found' }
                : { state:'pass', why:f.bankType }
        },
        {
            id: 'disclosure',
            label: 'Large-amount disclosure',
            run: f => {
                const amt = parseFloat((f.amount || '').replace(/[$,]/g, ''));
                if (isNaN(amt)) return { state:'skip', why:'amount not found' };
                if (amt < DISCLOSURE_LIMIT)
                    return { state:'pass', why:`under $${DISCLOSURE_LIMIT.toLocaleString()}` };
                return { state:'fail',
                         why:`$${amt.toLocaleString()} — confirm this is disclosed on a subsequent form` };
            }
        },
        {
            label: 'Contribution within annual limit',
            run: f => {
                const type = (f.acctType || '').toLowerCase();
                if (!LIMITED_TYPES.some(t => type === t))
                    return { state:'skip', why:`account type "${f.acctType || '?'}" — rule N/A` };
                const yr = parseInt(f.taxYear, 10);
                if (!f.taxYear) return { state:'fail', why:'tax year is blank on the ticket' };
                const lim = LIMITS[yr];
                if (!lim) return { state:'fail', why:`no limits configured for tax year ${yr}` };
                const age = parseFloat(f.age);
                if (isNaN(age)) return { state:'fail', why:'client age not found' };
                const cap = age >= CATCH_UP_AGE ? lim.catchUp : lim.base;
                const amt = parseFloat((f.amount || '').replace(/[$,]/g, ''));
                if (isNaN(amt)) return { state:'fail', why:'amount not found' };
                return amt <= cap
                    ? { state:'pass', why:`$${amt.toLocaleString()} ≤ $${cap.toLocaleString()} (age ${age}, TY ${yr})` }
                    : { state:'fail', why:`$${amt.toLocaleString()} exceeds $${cap.toLocaleString()} (age ${age}, TY ${yr})` };
            }
        },
    ];

    // ---------- fetching the bytes ---------------------------------------
    function findPdfStart(buf) {
        const head = new Uint8Array(buf.slice(0, 4096));
        const sig = [0x25, 0x50, 0x44, 0x46, 0x2d]; // %PDF-
        outer: for (let i = 0; i < head.length - 5; i++) {
            for (let j = 0; j < 5; j++) if (head[i + j] !== sig[j]) continue outer;
            return i;
        }
        return -1;
    }

    async function fetchOne(url) {
        const res = await fetch(url, {
            credentials: 'include',
            headers: { 'Accept': 'application/pdf,*/*' },
            referrer: location.href
        });
        if (!res.ok) throw new Error(`fetch ${res.status} on ${url}`);
        return { buf: await res.arrayBuffer(), type: res.headers.get('content-type') || '' };
    }

    async function fetchPDF(url, docID, d) {
        // Primary: the endpoint the wrapper frames.
        const direct = `${location.origin}/common/loadDoc_open.cfm?docID=${docID}&jobID=0`;
        try {
            const r = await fetchOne(direct);
            const at = findPdfStart(r.buf);
            if (at >= 0) return at ? r.buf.slice(at) : r.buf;
        } catch (e) {
            log('direct endpoint failed:', e.message);
        }

        // Fallback: fetch the wrapper and follow whatever it frames.
        let { buf, type } = await fetchOne(url);
        let at = findPdfStart(buf);
        if (at >= 0) return at ? buf.slice(at) : buf;

        const html = new TextDecoder().decode(buf.slice(0, 20000));
        log('wrapper HTML:\n' + html);

        const skip = /\.(css|js|png|gif|jpe?g|ico|svg|woff2?)(\?|$)/i;
        const cands = [];
        let mm;
        const attrRe = /(?:src|href|data|action)\s*=\s*["']([^"']+)["']/gi;
        while ((mm = attrRe.exec(html))) cands.push(mm[1]);
        const jsRe = /(?:location(?:\.href)?\s*=|window\.open\s*\(|url\s*:\s*)\s*["']([^"']+)["']/gi;
        while ((mm = jsRe.exec(html))) cands.push(mm[1]);

        const tries = [...new Set(cands)]
            .map(u => u.replace(/&amp;/g, '&'))
            .filter(u => u && !skip.test(u) && !/^(#|javascript:|mailto:)/i.test(u));

        for (const cand of tries) {
            const nested = new URL(cand, url).href;
            try {
                const second = await fetchOne(nested);
                const at2 = findPdfStart(second.buf);
                if (at2 >= 0) return at2 ? second.buf.slice(at2) : second.buf;
            } catch (e) { /* try the next one */ }
        }

        // Nothing worked — surface what the server actually sent.
        const status = d.getElementById('status');
        status.style.background = '#fee2e2';
        status.textContent = `Server returned ${type || 'unknown type'}, ` +
                             `${buf.byteLength} bytes — not a PDF. First 600 chars:`;
        const pre = d.createElement('pre');
        pre.style.cssText = 'white-space:pre-wrap;font:10px ui-monospace,monospace;' +
                            'max-height:340px;overflow:auto;margin-top:8px';
        pre.textContent = html.slice(0, 600);
        status.appendChild(pre);
        throw new Error('response was not a PDF');
    }

    // ---------- render ---------------------------------------------------
    const MARKS = { pass:'✓', fail:'✗', skip:'–' };

    async function run(win, pdfURL, docID, oes) {
        const d = win.document;
        const status = d.getElementById('status');

        const buf = await fetchPDF(pdfURL, docID, d);

        const items = await extractItems(buf);
        const { fields, rows } = parseFields(items);

        d.getElementById('sub').textContent = `docID ${docID} · ${items.length} text items`;
        status.style.display = 'none';
        d.getElementById('fields').style.display = 'block';

        const set = (id, val) => { d.getElementById(id).value = val || ''; };
        set('f_oes', oes);
        set('f_name', formatName(fields.name));
        set('f_rn', fields.routing);
        set('f_an', fields.bankAcct);
        set('f_type', (fields.bankType || '').replace(/^./, c => c.toUpperCase()));
        set('f_amt', fields.amount);

        // checks
        const box = d.getElementById('checks');
        const results = RULES.map(rule => ({ rule, r: rule.run(fields) }));
        for (const { rule, r } of results) {
            const div = d.createElement('div');
            div.className = 'check ' + r.state;
            div.innerHTML = `<span class="mark">${MARKS[r.state]}</span>` +
                            `<span class="body">${rule.label}` +
                            (r.why ? `<div class="why">${r.why}</div>` : '') + `</span>`;
            box.appendChild(div);
        }

        // A failed check blocks the copy. The disclosure rule is different:
        // it can't be satisfied from the ticket, only acknowledged.
        const hardFails = results.filter(x => x.r.state === 'fail' && x.rule.id !== 'disclosure');
        const needsAck  = results.some(x => x.rule.id === 'disclosure' && x.r.state === 'fail');

        const copyBtn  = d.getElementById('copyBtn');
        const blockMsg = d.getElementById('blockMsg');
        const ackBox   = d.getElementById('ackBox');

        if (needsAck) d.getElementById('bigWarn').style.display = 'block';
        if (hardFails.length) {
            blockMsg.style.display = 'block';
            blockMsg.textContent = 'Copy blocked — failed: ' +
                hardFails.map(x => x.rule.label).join(', ');
        }

        const refreshCopy = () => {
            const blocked = hardFails.length > 0 || (needsAck && !ackBox.checked);
            copyBtn.disabled = blocked;
            copyBtn.textContent = blocked ? 'Copy blocked' : 'Copy to clipboard';
        };
        ackBox.addEventListener('change', refreshCopy);
        refreshCopy();

        // output string, live-updates if a field is edited
        const out = d.getElementById('out');
        const build = () => {
            const g = id => d.getElementById(id).value.trim();
            out.value = `OES # ${g('f_oes')} ${g('f_name')} RN ${g('f_rn')} ` +
                        `AN ${g('f_an')} ${g('f_type')} ${g('f_amt')}`;
        };
        ['f_oes','f_name','f_rn','f_an','f_type','f_amt']
            .forEach(id => d.getElementById(id).addEventListener('input', build));
        build();

        copyBtn.addEventListener('click', () => {
            if (copyBtn.disabled) return;
            out.removeAttribute('readonly');
            out.focus(); out.select();
            let ok = false;
            try { ok = d.execCommand('copy'); } catch (e) {}
            out.setAttribute('readonly', '');
            if (!ok && win.navigator.clipboard) win.navigator.clipboard.writeText(out.value);
            copyBtn.textContent = 'Copied';
            setTimeout(refreshCopy, 1200);
        });

        // debug dump — coordinates for every text item, for tuning the parser
        const dump = d.getElementById('dump');
        dump.textContent = rows.map(r =>
            `p${r.page} y=${r.y.toFixed(1).padStart(7)}  ` +
            r.items.map(i => `[x=${i.x.toFixed(0)}] ${i.str}`).join('  |  ')
        ).join('\n');
        d.getElementById('dumpBtn').addEventListener('click', () => {
            const open = dump.style.display === 'block';
            dump.style.display = open ? 'none' : 'block';
            d.getElementById('dumpBtn').textContent = open ? 'Show extracted text' : 'Hide extracted text';
        });

        const noteEl = d.getElementById('notes');
        const note = sectionFiveNotes(rows);
        noteEl.textContent = note || '— no notes in Section Five —';
        noteEl.classList.toggle('empty', !note);

        log('fields', fields, 'notes:', note);
    }
})();