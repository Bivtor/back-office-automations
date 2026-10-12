// ==UserScript==
// @name         DQ Phase A  PERSH Scanner
// @namespace    http://tampermonkey.net/
// @version      1.2
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

    // Only activate when ?dq=1 is in the URL
    const params = new URLSearchParams(window.location.search);
    if (params.get('dq') !== '1') return;

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
            if (products.length === 0) return null;

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

        // Step 2: read product type from this page
        const productType = readProductType();
        const roId = getCurrentRoId();

        if (!roId) {
            log('ERROR: Could not read RO ID from URL');
            setStatus('#e63946');
            return;
        }

        if (productType === null) {
            log(`ERROR: Could not read product type for RO ${roId}`);
            setStatus('#e63946');
            return;
        }

        const isPersh = /persh/i.test(productType);
        $id('dq-current').textContent = `RO ID: ${roId}`;
        $id('dq-product').textContent = productType;
        $id('dq-product').style.color = isPersh ? '#e63946' : '#57cc99';

        log(`RO ${roId}  →  ${productType}${isPersh ? '  ⚑ PERSH' : ''}`);

        // Step 3: post result to server
        let resultData;
        try {
            resultData = await serverPost('/result', { roId, productType });
        } catch (e) {
            log('ERROR posting result to server: ' + e);
            setStatus('#e63946');
            return;
        }

        const newIndex = resultData.index ?? (pingData.index + 1);
        const total = resultData.total ?? pingData.total;
        setProgress(newIndex, total);

        // Step 4: navigate to next URL or finish
        if (resultData.done) {
            setStatus('#57cc99');
            $id('dq-progress-label').textContent = 'Scan complete!';
            $id('dq-current').textContent = 'All rows scanned.';
            $id('dq-product').textContent = '';
            log('All rows scanned. Assignment stage launching automatically.');
        } else {
            log(`Navigating to RO ${newIndex + 1} / ${total}...`);
            setTimeout(() => {
                window.location.href = resultData.nextUrl;
            }, 600);
        }
    }

    // Kick off after page settles
    window.addEventListener('load', () => setTimeout(run, 500));

})();
