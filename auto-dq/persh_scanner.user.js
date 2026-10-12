// ==UserScript==
// @name         DQ Phase A  PERSH Scanner
// @namespace    http://tampermonkey.net/
// @version      1.3
// @description  Reads product type from rolloverEdit.cfm and posts back to local PS server
// @author       Victor
// @match        https://www.partners.example.com/SimpleAlloc/Maintenance/rolloverEdit.cfm*
// @grant        GM_xmlhttpRequest
// @connect      localhost
// ==/UserScript==

(function () {
    'use strict';

    // Port 52731 is the dedicated Phase A port bound by server.ps1.
    // The GUI runs on 8080; both ports share the same listener and endpoints.
    const SERVER_PORT = 52731;
    const SERVER = `http://localhost:${SERVER_PORT}`;

    // Label posted when the Product column exists but every leg is empty.
    // Does not match /persh/i, so run.ps1 treats it as non-PERSH.
    const BLANK_LABEL = 'BLANK PRODUCT CODE';

    // Only activate when ?dq=1 is in the URL
    const params = new URLSearchParams(window.location.search);
    if (params.get('dq') !== '1') return;

    // Run-once guard: Tampermonkey can inject twice (e.g. iframes, re-injection),
    // and we never want two scans of the same page racing each other.
    if (window.__dqPhaseARan) return;
    window.__dqPhaseARan = true;

    // ── UI Pane ───────────────────────────────────────────────────────────────
    const pane = document.createElement('div');
    pane.id = 'dq-pane';
    Object.assign(pane.style, {
        position:     'fixed',
        bottom:       '16px',
        right:        '16px',
        width:        '320px',
        background:   '#1a1a2e',
        color:        '#e0e0e0',
        fontFamily:   'monospace',
        fontSize:     '12px',
        borderRadius: '8px',
        boxShadow:    '0 4px 20px rgba(0,0,0,0.5)',
        zIndex:       '999999',
        overflow:     'hidden',
        userSelect:   'none',
    });

    pane.innerHTML = `
        <div id="dq-header" style="
            background: #16213e;
            padding: 8px 12px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            cursor: move;
            border-bottom: 1px solid #0f3460;
        ">
            <span style="color:#00b4d8; font-weight:bold;">DQ Phase A  PERSH Scan</span>
            <span id="dq-status-dot" style="
                width:10px; height:10px; border-radius:50%;
                background:#888; display:inline-block;
            "></span>
        </div>
        <div style="padding: 10px 12px;">
            <div id="dq-progress-label" style="margin-bottom:6px; color:#aaa;">Connecting to server...</div>
            <div style="background:#0f3460; border-radius:4px; height:8px; overflow:hidden; margin-bottom:10px;">
                <div id="dq-progress-bar" style="height:100%; width:0%; background:#00b4d8; transition:width 0.3s;"></div>
            </div>
            <div id="dq-current" style="color:#e0e0e0; margin-bottom:4px;"></div>
            <div id="dq-product" style="font-size:14px; font-weight:bold; margin-bottom:8px;"></div>
            <a id="dq-jump" href="#" style="
                display:none; margin-bottom:8px; padding:4px 8px;
                background:#0f3460; color:#00b4d8; border-radius:4px;
                text-decoration:none; text-align:center;
            "></a>
            <div id="dq-log" style="
                max-height: 120px;
                overflow-y: auto;
                border-top: 1px solid #0f3460;
                padding-top: 6px;
                color: #888;
                font-size: 11px;
                line-height: 1.6;
            "></div>
        </div>
    `;
    document.body.appendChild(pane);

    // ── Draggable ─────────────────────────────────────────────────────────────
    const header = document.getElementById('dq-header');
    let dragging = false, dragX = 0, dragY = 0;
    header.addEventListener('mousedown', e => {
        dragging = true;
        dragX = e.clientX - pane.getBoundingClientRect().left;
        dragY = e.clientY - pane.getBoundingClientRect().top;
    });
    document.addEventListener('mousemove', e => {
        if (!dragging) return;
        pane.style.left = (e.clientX - dragX) + 'px';
        pane.style.top = (e.clientY - dragY) + 'px';
        pane.style.right = 'auto';
        pane.style.bottom = 'auto';
    });
    document.addEventListener('mouseup', () => dragging = false);

    // ── Helpers ───────────────────────────────────────────────────────────────
    const $id = id => document.getElementById(id);
    const log = msg => {
        const el = $id('dq-log');
        el.innerHTML += `<div>${msg}</div>`;
        el.scrollTop = el.scrollHeight;
    };

    function setStatus(color) {
        $id('dq-status-dot').style.background = color;
    }

    function setProgress(index, total) {
        const pct = total > 0 ? Math.round((index / total) * 100) : 0;
        $id('dq-progress-bar').style.width = pct + '%';
        $id('dq-progress-label').textContent = `${index} / ${total}  (${pct}%)`;
    }

    // Manual "go to the RO the server is waiting on" link. Used instead of an
    // auto-redirect when this page is NOT the expected RO, so a stray second tab
    // can't start ping-ponging through the queue alongside the main one.
    function showJump(url, label) {
        if (!url) return;
        const a = $id('dq-jump');
        a.href = url;
        a.textContent = label;
        a.style.display = 'block';
    }

    function serverGet(path) {
        return new Promise((resolve, reject) => {
            GM_xmlhttpRequest({
                method: 'GET',
                url:    `${SERVER}${path}`,
                onload: r => {
                    try { resolve(JSON.parse(r.responseText)); }
                    catch(e) { reject(e); }
                },
                onerror: reject,
            });
        });
    }

    function serverPost(path, data) {
        return new Promise((resolve, reject) => {
            GM_xmlhttpRequest({
                method:  'POST',
                url:     `${SERVER}${path}`,
                headers: { 'Content-Type': 'application/json' },
                data:    JSON.stringify(data),
                onload:  r => {
                    try { resolve(JSON.parse(r.responseText)); }
                    catch(e) { reject(e); }
                },
                onerror: reject,
            });
        });
    }

    // ── Read product type from the page ───────────────────────────────────────
    // The split-accounts table lives INSIDE <form name="form1">, so a
    // "body > table" path never reaches it. Instead we anchor on the
    // "Product" column header, find its index, and read the data row(s)
    // beneath it. This survives the form wrapper, the foster-parented
    // hidden <input> rows, and minor layout drift.
    //
    // Layout:  Account | Account Type | Product | Target Company | ...
    //                                   ^^^^^^^ this column
    //
    // Returns:
    //   "NMZZ" / "A; B"     -> product code(s) found
    //   BLANK_LABEL          -> Product column found, but every leg is empty
    //   null                 -> Product column NOT found (page didn't render,
    //                           layout changed, error page) -- a real failure
    function readProductType() {
        try {
            // 1. Locate the "Product" column header cell (exact text match).
            const headerTd = Array.from(document.querySelectorAll('td'))
            .find(td => td.textContent.replace(/\u00a0/g, ' ').trim().toLowerCase() === 'product');
            if (!headerTd) return null;

            const headerRow = headerTd.closest('tr');
            const table = headerTd.closest('table');
            if (!headerRow || !table) return null;

            // 2. Column index of "Product" among the header row's <td>s.
            const headerTds = Array.from(headerRow.children).filter(el => el.tagName === 'TD');
            const colIdx = headerTds.indexOf(headerTd);
            if (colIdx < 0) return null;

            // 3. Read every data row after the header within this same table.
            //    A split rollover can have multiple legs; if ANY leg is a
            //    Pershing product we want it to surface, so collect them all.
            const allRows = Array.from(table.querySelectorAll('tr'));
            const startIdx = allRows.indexOf(headerRow) + 1;
            const products = [];
            for (let i = startIdx; i < allRows.length; i++) {
                const tds = Array.from(allRows[i].children).filter(el => el.tagName === 'TD');
                if (tds.length <= colIdx) continue; // header sub-rows / spacers
                const val = tds[colIdx].textContent.replace(/\u00a0/g, ' ').trim();
                if (val) products.push(val);
            }

            // Header is present, so the page rendered -- an empty column is a
            // legitimate data state, not a read failure.
            if (products.length === 0) return BLANK_LABEL;

            // Single leg -> plain value (e.g. "NMZZ"); multi-leg -> joined.
            return products.join('; ');
        } catch (e) {
            return null;
        }
    }

    // ── Get RO ID from the current URL ────────────────────────────────────────
    function getCurrentRoId() {
        return new URLSearchParams(window.location.search).get('ID');
    }

    // ── Main flow ─────────────────────────────────────────────────────────────
    let posted = false; // second guard: never POST twice from this page load

    async function run() {
        setStatus('#888');
        $id('dq-progress-label').textContent = 'Connecting to server...';

        // Step 1: ping server -- confirms PhaseA stage is active
        let pingData;
        try {
            pingData = await serverGet('/ping');
        } catch (e) {
            setStatus('#e63946');
            $id('dq-progress-label').textContent = 'ERROR: Server not reachable';
            log('Could not reach PS server on port ' + SERVER_PORT);
            log('Make sure server.ps1 is running (START.bat).');
            return;
        }

        if (!pingData.ok) {
            setStatus('#e63946');
            // done=true + ok=false means server is up but not in PhaseA stage
            const note = pingData.note || '';
            log('Server responded but Phase A is not active.');
            if (note) log(note);
            $id('dq-progress-label').textContent = 'Not in Phase A — run from GUI first.';
            return;
        }

        if (pingData.done) {
            setStatus('#57cc99');
            $id('dq-progress-label').textContent = 'Scan already complete!';
            log('All rows already scanned — nothing to do.');
            return;
        }

        setStatus('#00b4d8');
        log(`Server connected. ${pingData.total} rows to scan.`);
        setProgress(pingData.index, pingData.total);

        const roId = getCurrentRoId();
        if (!roId) {
            log('ERROR: Could not read RO ID from URL');
            setStatus('#e63946');
            return;
        }
        $id('dq-current').textContent = `RO ID: ${roId}`;

        // Step 2: duplicate / out-of-order check BEFORE reading or posting.
        // Refresh, Back button, a reopened tab, or a second tab will land here
        // on an RO the server isn't waiting for. Don't post -- just offer a link
        // to the RO that IS expected.
        if (pingData.expectedRoId && String(pingData.expectedRoId) !== String(roId)) {
            setStatus('#f4a261');
            $id('dq-progress-label').textContent = 'Not the current RO — skipped.';
            log(`RO ${roId} is not next in queue (server expects ${pingData.expectedRoId}).`);
            log('Nothing posted. Scan state unchanged.');
            showJump(pingData.expectedUrl, `Go to RO ${pingData.expectedRoId} →`);
            return;
        }

        // Step 3: read product type from this page
        const productType = readProductType();

        if (productType === null) {
            log(`ERROR: Could not find Product column for RO ${roId}`);
            log('Page may not have loaded — refresh to retry.');
            setStatus('#e63946');
            return;
        }

        const isBlank = productType === BLANK_LABEL;
        const isPersh = /persh/i.test(productType);
        $id('dq-product').textContent = productType;
        $id('dq-product').style.color = isPersh ? '#e63946' : isBlank ? '#f4a261' : '#57cc99';

        log(`RO ${roId}  →  ${productType}${isPersh ? '  ⚑ PERSH' : ''}${isBlank ? '  ⚠ BLANK' : ''}`);

        // Step 4: post result to server (once per page load)
        if (posted) return;
        posted = true;

        let resultData;
        try {
            resultData = await serverPost('/result', { roId, productType });
        } catch (e) {
            log('ERROR posting result to server: ' + e);
            setStatus('#e63946');
            posted = false; // allow a manual refresh to retry
            return;
        }

        if (!resultData.ok) {
            setStatus('#e63946');
            log('Server rejected result: ' + (resultData.error || resultData.note || 'unknown'));
            return;
        }

        // Server-side backstop: it already had this RO (e.g. two tabs posted
        // the same page). Nothing was recorded or advanced.
        if (resultData.accepted === false) {
            setStatus('#f4a261');
            log(`Server ignored RO ${roId} (${resultData.reason || 'duplicate'}). Scan state unchanged.`);
            setProgress(resultData.index, resultData.total);
            if (resultData.done) {
                $id('dq-progress-label').textContent = 'Scan already complete!';
            } else {
                showJump(resultData.nextUrl, 'Go to next RO in queue →');
            }
            return;
        }

        const newIndex = resultData.index ?? (pingData.index + 1);
        const total = resultData.total ?? pingData.total;
        setProgress(newIndex, total);

        // Step 5: navigate to next URL or finish
        if (resultData.done) {
            setStatus('#57cc99');
            $id('dq-progress-label').textContent = 'Scan complete!';
            $id('dq-current').textContent = 'All rows scanned.';
            $id('dq-product').textContent = '';
            log('All rows scanned. Assignment stage launching automatically.');
        } else {
            log(`Navigating to RO ${newIndex + 1} / ${total}...`);
            setTimeout(() => {
                // replace() keeps already-scanned ROs out of Back-button history
                window.location.replace(resultData.nextUrl);
            }, 600);
        }
    }

    // Kick off after page settles
    if (document.readyState === 'complete') {
        setTimeout(run, 500);
    } else {
        window.addEventListener('load', () => setTimeout(run, 500), { once: true });
    }

})();