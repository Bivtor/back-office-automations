// ==UserScript==
// @name         UMB ACH — Template Find & Verify
// @namespace    http://tampermonkey.net/
// @version      1.0
// @description  Finds a PPD template from a Trade Ticket payload, then verifies RN/AN on the payment page before filling the amount
// @match        https://umbdirect.umb.com/rwd-web/*
// @grant        none
// @noframes
// @run-at       document-idle
// ==/UserScript==

// No @require, no GM_* storage, nothing persisted, nothing added to any URL.
// State lives in the closure below and dies with the tab.

(function () {
    'use strict';

    // ===== CONFIG ========================================================
    const TYPE_DELAY   = 25;      // ms per simulated keystroke
    const REQUIRE_PPD  = true;    // only match PPD templates
    const FILL_AMOUNT  = true;    // false -> verify only, type the amount yourself
    // The template search grid ('create payment from template'):
    const TEMPLATE_SEARCH_URL = 'https://umbdirect.umb.com/rwd-web/main/payments/createACHFromTemplate';
    // =====================================================================

    let pending = null;           // payload carried across the route change
    const sleep = ms => new Promise(r => setTimeout(r, ms));

    // ---------- payload ---------------------------------------------------
    const PAYLOAD_RE = new RegExp(
        'OES\\s*#\\s*(\\d+)\\s+' + '(.+?)\\s+' + 'RN\\s*(\\d{9})\\s+' +
        'AN\\s*([\\d\\-]+)\\s+' + '(checking|savings)\\s+' + '\\$?([\\d,]+\\.\\d{2})', 'i');

    function parsePayload(text) {
        const m = (text || '').replace(/\s+/g, ' ').trim().match(PAYLOAD_RE);
        if (!m) return null;
        return {
            oes: m[1], name: m[2].trim(), routing: m[3], account: m[4],
            type: m[5].replace(/^./, c => c.toUpperCase()), amount: m[6]
        };
    }

    // ---------- shared helpers -------------------------------------------
    function waitFor(sel, { timeout = 15000, root = document, test } = {}) {
        return new Promise((resolve, reject) => {
            const t0 = Date.now();
            (function poll() {
                const el = root.querySelector(sel);
                if (el && (!test || test(el))) return resolve(el);
                if (Date.now() - t0 > timeout) return reject(new Error('timed out: ' + sel));
                setTimeout(poll, 100);
            })();
        });
    }

    function setNativeValue(el, value) {
        const d = Object.getOwnPropertyDescriptor(Object.getPrototypeOf(el), 'value');
        if (d && d.set) d.set.call(el, value); else el.value = value;
    }

    async function typeInto(el, text) {
        el.focus();
        setNativeValue(el, '');
        el.dispatchEvent(new Event('input', { bubbles: true }));
        for (const ch of String(text)) {
            const max = parseInt(el.getAttribute('maxlength') || '0', 10);
            if (max > 0 && el.value.length >= max) break;
            el.dispatchEvent(new KeyboardEvent('keydown', { key: ch, bubbles: true }));
            setNativeValue(el, el.value + ch);
            el.dispatchEvent(new InputEvent('input', { bubbles: true, data: ch, inputType: 'insertText' }));
            el.dispatchEvent(new KeyboardEvent('keyup', { key: ch, bubbles: true }));
            await sleep(TYPE_DELAY);
        }
        el.dispatchEvent(new Event('change', { bubbles: true }));
        return el.value;
    }

    function clickIt(el) {
        const o = { bubbles: true, cancelable: true, view: window };
        el.dispatchEvent(new MouseEvent('mouseover', o));
        try { el.dispatchEvent(new PointerEvent('pointerdown', o)); } catch (e) {}
        el.dispatchEvent(new MouseEvent('mousedown', o));
        if (el.focus) el.focus();
        try { el.dispatchEvent(new PointerEvent('pointerup', o)); } catch (e) {}
        el.dispatchEvent(new MouseEvent('mouseup', o));
        el.click();
    }

    const norm = s => (s || '').toLowerCase().replace(/\s+/g, ' ').trim();
    const digits = s => (s || '').replace(/\D/g, '');

    // ---------- template name matching ------------------------------------
    // Grid truncates long names ("Smith, Kathleen D'Amat..."), and duplicate
    // accounts get a numeric suffix ("Adams, Anna 0611").
    function nameMatch(rowName, target) {
        const r = norm(rowName), t = norm(target);
        if (!r) return null;
        if (r === t) return { how: 'exact', suffix: '' };
        const trunc = r.replace(/(\.{3}|…)\s*$/, '');
        if (trunc !== r && t.startsWith(trunc.trim())) return { how: 'truncated', suffix: '' };
        const suffixed = r.match(/^(.*?)\s+(\d{3,5})$/);
        if (suffixed && norm(suffixed[1]) === t)
            return { how: 'suffix ' + suffixed[2], suffix: suffixed[2] };
        return null;
    }

    // A duplicate template's numeric suffix is the last 4 of its account number,
    // so the payload's account picks the right one. Returns {row, why} or null.
    function chooseMatch(matches, p) {
        const last4 = digits(p.account).slice(-4);
        if (matches.length === 1)
            return { row: matches[0], why: 'only candidate' };

        const bySuffix = matches.filter(m => m.suffix && m.suffix.slice(-4) === last4);
        if (bySuffix.length === 1)
            return { row: bySuffix[0], why: `suffix ${bySuffix[0].suffix} matches account ...${last4}` };

        const plain = matches.filter(m => !m.suffix);
        if (!bySuffix.length && plain.length === 1)
            return { row: plain[0], why: `no template ends in ${last4} — falling back to the unsuffixed one` };

        return null;
    }

    function gridRows() {
        const table = document.querySelector('#templateCenterGrid-table');
        if (!table) return [];
        return [...table.querySelectorAll('tbody tr')].map((tr, i) => {
            const cell = n => tr.querySelector(`td[headers^="templateCenterGrid-header-header-${n}"]`);
            const nameCell = cell(1);
            return {
                tr, index: i,
                name: nameCell ? (nameCell.querySelector('.cell-primary-data')?.innerText || '').trim() : '',
                desc: nameCell ? (nameCell.querySelector('.cell-secondary-data')?.innerText || '').trim() : '',
                type: (cell(4)?.innerText || '').trim(),
                checkbox: tr.querySelector('input.checkinput[type="checkbox"]')
            };
        }).filter(r => r.name);
    }

    // ---------- search page ------------------------------------------------
    async function findTemplate(p, step, onMatches) {
        const input = await waitFor('#advancedSearchTemplateNameReference');
        const surname = p.name.split(',')[0].trim();
        await typeInto(input, surname + '*');
        step('Search', 'ok', surname + '*');

        clickIt(await waitFor('#btnSearch'));
        await sleep(1200);
        await waitFor('#templateCenterGrid-table tbody tr', { timeout: 20000 });
        await sleep(400);

        const rows = gridRows().filter(r => !REQUIRE_PPD || /^PPD/i.test(r.type));
        const matches = rows
            .map(r => ({ ...r, ...(nameMatch(r.name, p.name) || {}) }))
            .filter(r => r.how);

        if (!matches.length) {
            step('Match', 'fail', `no PPD template matching "${p.name}" on this page`);
            return;
        }
        step('Match', 'ok', `${matches.length} candidate${matches.length > 1 ? 's' : ''}`);
        onMatches(matches);
    }

    async function openPayment(row, p, step) {
        if (!row.checkbox) { step('Select row', 'fail', 'no checkbox found'); return; }
        if (!row.checkbox.checked) clickIt(row.checkbox);
        await sleep(400);
        step('Select row', row.checkbox.checked ? 'ok' : 'fail', row.name);
        if (!row.checkbox.checked) return;

        const btn = await waitFor('#btnCreatePayment', { test: b => !b.disabled });
        pending = p;                       // carried across the route change
        clickIt(btn);
        step('Create Payment', 'ok', 'opening…');
    }

    // ---------- payment page ----------------------------------------------
    function readPaymentPage() {
        const cell = n => document.querySelector(
            `#recipients-table tbody td[headers^="recipients-header-header-${n}"]`);
        const templateName = [...document.querySelectorAll('ol-preview-data')]
            .map(e => e.innerText.trim())
            .find(t => /^Template Name/i.test(t))?.replace(/^Template Name\s*/i, '').trim() || '';
        return {
            templateName,
            bankId:  (cell(1)?.querySelector('.cell-primary-data')?.innerText || '').trim(),
            bankName:(cell(1)?.querySelector('.cell-secondary-data')?.innerText || '').trim(),
            account: (cell(2)?.querySelector('.cell-primary-data')?.innerText || '').trim(),
            acctType:(cell(2)?.querySelector('.cell-secondary-data')?.innerText || '').trim()
        };
    }

    async function verifyAndFill(p, step) {
        await waitFor('#recipients-table tbody td[headers^="recipients-header-header-1"]');
        const page = readPaymentPage();
        let ok = true;

        // routing
        if (digits(page.bankId) === digits(p.routing)) {
            step('Routing number', 'ok', page.bankId);
        } else {
            ok = false;
            step('Routing number', 'fail', `ticket ${p.routing} · template ${page.bankId || '(none)'}`);
        }

        // account — full compare when shown in full, last-4 when masked
        const masked = /[*x•]/i.test(page.account);
        if (!masked && digits(page.account) === digits(p.account)) {
            step('Account number', 'ok', page.account);
        } else if (masked && digits(page.account) &&
                   digits(p.account).endsWith(digits(page.account).slice(-4))) {
            step('Account number', 'ok', `last 4 only (${page.account}) — full number not shown`);
        } else {
            ok = false;
            step('Account number', 'fail', `ticket ${p.account} · template ${page.account || '(none)'}`);
        }

        // account type
        if (norm(page.acctType) === norm(p.type)) {
            step('Account type', 'ok', page.acctType);
        } else {
            ok = false;
            step('Account type', 'fail', `ticket ${p.type} · template ${page.acctType || '(none)'}`);
        }

        step('Template', 'skip', `${page.templateName} · ${page.bankName}`);

        if (!ok) {
            step('STOPPED', 'fail', 'mismatch — amount not filled. Check the ticket against the template.');
            return false;
        }

        if (!FILL_AMOUNT) {
            step('Amount', 'skip', `verified — type ${p.amount} yourself`);
            return true;
        }
        const amt = await waitFor('#recipientAmount_0_input');
        step('Amount', 'ok', await typeInto(amt, p.amount.replace(/,/g, '')));

        const cont = document.querySelector('#btnContinue');
        if (cont) cont.focus();
        step('Continue', 'skip', 'focused — review, then continue and enter your 2FA code');
        return true;
    }

    // ---------- panel ------------------------------------------------------
    const PANEL_ID = 'tt-umb-panel';
    const MARK = { ok: ['✓', '#16a34a'], warn: ['!', '#d97706'],
                   skip: ['–', '#a1a1aa'], fail: ['✗', '#dc2626'] };

    function buildPanel(mode) {
        let wrap = document.getElementById(PANEL_ID);
        if (wrap && wrap.dataset.mode === mode) return;
        if (wrap) wrap.remove();

        wrap = document.createElement('div');
        wrap.id = PANEL_ID;
        wrap.dataset.mode = mode;
        wrap.style.cssText = `position:fixed; right:16px; bottom:16px; width:340px;
            z-index:2147483000; background:#fff; border:1px solid #d4d4d8; border-radius:6px;
            box-shadow:0 4px 16px rgba(0,0,0,.18); color:#18181b;
            font:13px/1.45 -apple-system,Segoe UI,Roboto,sans-serif;`;

        const head = document.createElement('div');
        head.style.cssText = `padding:9px 12px; border-bottom:1px solid #e4e4e7; font-weight:600;
            display:flex; justify-content:space-between; align-items:center;`;
        head.textContent = mode === 'search' ? 'Find template' : 'Verify payment';
        const hide = document.createElement('button');
        hide.textContent = '–';
        hide.style.cssText = 'border:0;background:none;font-size:17px;cursor:pointer;line-height:1;';
        head.appendChild(hide);

        const body = document.createElement('div');
        body.style.cssText = 'padding:12px;';
        const logBox = document.createElement('div');
        logBox.style.cssText = 'max-height:300px; overflow:auto;';

        hide.addEventListener('click', () => {
            const shown = body.style.display !== 'none';
            body.style.display = shown ? 'none' : 'block';
            hide.textContent = shown ? '+' : '–';
        });

        const step = (label, state, detail) => {
            const [g, c] = MARK[state] || MARK.skip;
            const row = document.createElement('div');
            row.style.cssText = 'display:flex; gap:7px; margin-bottom:5px; font-size:12px;';
            row.innerHTML = `<span style="color:${c};font-weight:700">${g}</span>` +
                `<span><b>${label}</b>` +
                (detail ? `<div style="color:#71717a;font-size:11px">${detail}</div>` : '') + `</span>`;
            logBox.appendChild(row);
            logBox.scrollTop = logBox.scrollHeight;
        };

        wrap.append(head, body);
        document.body.appendChild(wrap);
        return { wrap, body, logBox, step };
    }

    function pasteBox(body, onGo) {
        const ta = document.createElement('textarea');
        ta.placeholder = 'Paste the payload here…';
        ta.style.cssText = `width:100%; height:58px; padding:7px; resize:vertical;
            border:1px solid #d4d4d8; border-radius:4px; font:12px ui-monospace,Menlo,monospace;`;
        const go = document.createElement('button');
        go.textContent = 'Go';
        go.style.cssText = `width:100%; margin:8px 0; padding:9px; border:0; border-radius:5px;
            background:#18181b; color:#fff; font-size:13px; cursor:pointer;`;
        go.addEventListener('click', () => onGo(ta.value, go));
        body.append(ta, go);
        return ta;
    }

    function renderSearchPanel() {
        const ui = buildPanel('search');
        if (!ui) return;
        const { body, logBox, step } = ui;

        const picker = document.createElement('div');
        picker.style.cssText = 'margin-bottom:8px;';

        pasteBox(body, async (text, go) => {
            logBox.textContent = '';
            picker.textContent = '';
            const p = parsePayload(text);
            if (!p) { step('Payload', 'fail', "couldn't parse that"); return; }
            step('Payload', 'ok', `${p.name} · ${p.type} · $${p.amount}`);
            go.disabled = true;
            try {
                await findTemplate(p, step, (matches) => {
                    const choice = chooseMatch(matches, p);
                    if (choice) {
                        step('Chose', choice.row.suffix ? 'ok' : 'warn',
                             `${choice.row.name} — ${choice.why}`);
                        openPayment(choice.row, p, step);
                        return;
                    }
                    step('Pick one', 'warn',
                         `can't tell which account — ticket ends ...${digits(p.account).slice(-4)}`);
                    matches.forEach(m => {
                        const b = document.createElement('button');
                        b.textContent = `${m.name}  (${m.how})`;
                        b.style.cssText = `display:block; width:100%; text-align:left; margin-bottom:4px;
                            padding:7px; border:1px solid #d4d4d8; border-radius:4px;
                            background:#fafafa; cursor:pointer; font-size:12px;`;
                        b.addEventListener('click', () => openPayment(m, p, step));
                        picker.appendChild(b);
                    });
                });
            } catch (e) {
                step('Stopped', 'fail', e.message);
            } finally { go.disabled = false; }
        });

        body.append(picker, logBox);
    }

    function renderPaymentPanel() {
        const ui = buildPanel('payment');
        if (!ui) return;
        const { body, logBox, step } = ui;

        if (pending) {
            const p = pending;
            pending = null;                       // one payment per hand-off
            body.append(logBox);
            step('Payload', 'ok', `${p.name} · $${p.amount}`);
            verifyAndFill(p, step).catch(e => step('Stopped', 'fail', e.message));
        } else {
            pasteBox(body, async (text, go) => {
                logBox.textContent = '';
                const p = parsePayload(text);
                if (!p) { step('Payload', 'fail', "couldn't parse that"); return; }
                go.disabled = true;
                try { await verifyAndFill(p, step); }
                catch (e) { step('Stopped', 'fail', e.message); }
                finally { go.disabled = false; }
            });
            body.append(logBox);
        }
    }


    // ---------- back to template search ------------------------------------
    // The success screen's "Payment Center" button is the natural exit point,
    // so point it at the template search page instead.
    const BACK_LABEL = 'Back to Templates';

    function retargetSuccessButton() {
        const btn = document.querySelector('alert-component .success-action button');
        if (!btn || btn.dataset.ttBackPatched) return;
        btn.dataset.ttBackPatched = '1';

        const label = btn.querySelector('.button-label');
        if (label) label.textContent = BACK_LABEL;

        // Capture phase on the element itself, so this runs before Angular's
        // own click binding and stops it.
        btn.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopImmediatePropagation();
            location.href = TEMPLATE_SEARCH_URL;
        }, true);
    }

    // ---------- route watching --------------------------------------------
    function route() {
        if (document.querySelector('#advancedSearchTemplateNameReference')) return 'search';
        if (document.querySelector('#recipients-table') &&
            document.querySelector('#Create_PPD_Payment_from_Template')) return 'payment';
        return null;
    }

    function tick() {
        retargetSuccessButton();   // success screen can outlive the route
        const r = route();
        if (r === 'search') renderSearchPanel();
        else if (r === 'payment') renderPaymentPanel();
        else document.getElementById(PANEL_ID)?.remove();
    }

    new MutationObserver(tick).observe(document.body, { childList: true, subtree: true });
    tick();
})();