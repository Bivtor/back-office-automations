// ==UserScript==
// @name         UMB ACH Template — Fill from Trade Ticket
// @namespace    http://tampermonkey.net/
// @version      1.0
// @description  Paste a Trade Ticket payload and fill the ACH PPD template + recipient forms
// @match        https://umbdirect.umb.com/rwd-web/*
// @grant        none
// @noframes
// @run-at       document-idle
// ==/UserScript==

// NOTE: no @require, no GM_* storage, nothing persisted. Everything this script
// needs is inline, and the payload lives only in the page until you close it.

(function () {
    'use strict';

    // ===== CONFIG ========================================================
    const ENTRY_DESC       = 'Lockbox TT';   // Company Entry Description, every time
    const AUTO_CREATE_NEW  = true;           // click "Create New" to open the recipient modal
    const AUTO_CONTINUE    = false;          // click "Continue" at the end of the modal
    const TYPE_DELAY       = 25;             // ms between simulated keystrokes
    const STEP_PAUSE       = 220;            // ms between fields
    // The template search grid ('create payment from template'):
    const TEMPLATE_SEARCH_URL = 'https://umbdirect.umb.com/rwd-web/main/payments/createACHFromTemplate';
    // =====================================================================

    const sleep = ms => new Promise(r => setTimeout(r, ms));

    // ---------- payload ---------------------------------------------------
    // OES # 1234567 Doe, Jane RN 123456789 AN 987654321 Checking $8,600.00
    const PAYLOAD_RE = new RegExp(
        'OES\\s*#\\s*(\\d+)\\s+' +      // 1 OES
        '(.+?)\\s+' +                   // 2 name
        'RN\\s*(\\d{9})\\s+' +          // 3 routing
        'AN\\s*([\\d\\-]+)\\s+' +       // 4 account
        '(checking|savings)\\s+' +      // 5 type
        '\\$?([\\d,]+\\.\\d{2})',       // 6 amount
        'i');

    function parsePayload(text) {
        const m = (text || '').replace(/\s+/g, ' ').trim().match(PAYLOAD_RE);
        if (!m) return null;
        return {
            oes:     m[1],
            name:    m[2].trim(),
            routing: m[3],
            account: m[4],
            type:    m[5].replace(/^./, c => c.toUpperCase()),
            amount:  m[6]
        };
    }

    // ---------- DOM helpers ----------------------------------------------
    function waitFor(sel, { timeout = 10000, root = document, test } = {}) {
        return new Promise((resolve, reject) => {
            const started = Date.now();
            (function poll() {
                const el = root.querySelector(sel);
                if (el && (!test || test(el))) return resolve(el);
                if (Date.now() - started > timeout)
                    return reject(new Error('timed out waiting for ' + sel));
                setTimeout(poll, 100);
            })();
        });
    }

    const visible = el => !!el && !el.hidden && !!el.offsetParent;

    // Framework-safe value assignment (bypasses Angular's cached value).
    function setNativeValue(el, value) {
        const desc = Object.getOwnPropertyDescriptor(Object.getPrototypeOf(el), 'value');
        if (desc && desc.set) desc.set.call(el, value);
        else el.value = value;
    }

    // Character-by-character entry with real key events.
    async function typeInto(el, text) {
        el.focus();
        el.dispatchEvent(new FocusEvent('focus', { bubbles: true }));

        setNativeValue(el, '');
        el.dispatchEvent(new Event('input', { bubbles: true }));

        for (const ch of String(text)) {
            const max = parseInt(el.getAttribute('maxlength') || '0', 10);
            if (max > 0 && el.value.length >= max) break;   // field is full, stop cleanly

            el.dispatchEvent(new KeyboardEvent('keydown',  { key: ch, bubbles: true }));
            el.dispatchEvent(new KeyboardEvent('keypress', { key: ch, bubbles: true }));
            setNativeValue(el, el.value + ch);
            el.dispatchEvent(new InputEvent('input', {
                bubbles: true, data: ch, inputType: 'insertText'
            }));
            el.dispatchEvent(new KeyboardEvent('keyup', { key: ch, bubbles: true }));
            await sleep(TYPE_DELAY);
        }

        el.dispatchEvent(new Event('change', { bubbles: true }));
        el.blur();
        el.dispatchEvent(new FocusEvent('blur', { bubbles: true }));
        return el.value;
    }

    function clickIt(el) {
        const o = { bubbles: true, cancelable: true, view: window };
        el.dispatchEvent(new MouseEvent('mouseover', o));
        el.dispatchEvent(new MouseEvent('mouseenter', { bubbles: false }));
        el.dispatchEvent(new MouseEvent('mousemove', o));
        try { el.dispatchEvent(new PointerEvent('pointerdown', o)); } catch (e) {}
        el.dispatchEvent(new MouseEvent('mousedown', o));
        if (el.focus) el.focus();
        try { el.dispatchEvent(new PointerEvent('pointerup', o)); } catch (e) {}
        el.dispatchEvent(new MouseEvent('mouseup', o));
        el.click();
    }

    // The options list is portalled into <ol-overlay-holder> at document level,
    // NOT inside the ol-dropdown. Ids are predictable:
    //   #<base>_optionsList   and   #<base>_option_<optionValue>
    const optionsListFor = base => document.getElementById(base + '_optionsList');

    async function closeDropdown(base) {
        if (!optionsListFor(base)) return 'closed itself';

        const esc = { key: 'Escape', code: 'Escape', keyCode: 27, which: 27, bubbles: true };
        (optionsListFor(base) || document).dispatchEvent(new KeyboardEvent('keydown', esc));
        document.dispatchEvent(new KeyboardEvent('keydown', esc));
        await sleep(160);
        if (!optionsListFor(base)) return 'escape';

        // Click the dropdown overlay's OWN backdrop. Never document.body — that
        // dismisses the Create Recipient modal underneath.
        const backdrop = document.querySelector('.modal.dropdown-wrapper');
        if (backdrop) {
            clickIt(backdrop);
            await sleep(180);
            if (!optionsListFor(base)) return 'backdrop';
        }

        // Last resort — pull the portal out of the DOM.
        const wrapper = document.querySelector('ol-overlay-dropdwon-wrapper') ||
                        document.querySelector('ol-dropdown-options')?.closest('.modal') ;
        if (wrapper) wrapper.remove();
        return optionsListFor(base) ? 'still open' : 'removed';
    }

    async function selectDropdown(base, matcher) {
        const root = await waitFor('#' + base);
        await waitFor('select', {
            root,
            test: s => [...s.options].some(o => !o.classList.contains('placeholder'))
        });
        const nativeSel = root.querySelector('select');
        const opt = [...nativeSel.options].find(o =>
            !o.classList.contains('placeholder') && matcher(o.textContent.trim(), o.value));
        if (!opt) throw new Error('no matching option');

        const wanted = opt.textContent.trim();
        const display = () => (document.getElementById(base + '_displayText')?.textContent || '').trim();
        const settled = () => display() === wanted;

        // A — open the widget and click the portalled option
        const link = root.querySelector('a.dropdown-value-link');
        if (link) {
            clickIt(link);
            let list = null;
            try { list = await waitFor('#' + base + '_optionsList', { timeout: 3000 }); }
            catch (e) { /* portal never opened */ }

            if (list) {
                const choice = document.getElementById(`${base}_option_${opt.value}`) ||
                    [...list.querySelectorAll('[role="option"]')].find(e =>
                        !e.classList.contains('placeholder') &&
                        (e.getAttribute('aria-label') || e.textContent).trim() === wanted);

                if (choice) {
                    clickIt(choice);
                    await sleep(280);
                    const how = await closeDropdown(base);
                    if (settled()) return { via: 'widget', text: display(), closed: how };
                } else {
                    await closeDropdown(base);
                }
            }
        }

        // B — hidden native select
        nativeSel.focus();
        setNativeValue(nativeSel, opt.value);
        nativeSel.dispatchEvent(new Event('change', { bubbles: true }));
        nativeSel.dispatchEvent(new Event('input',  { bubbles: true }));
        nativeSel.blur();
        await sleep(250);
        const how = await closeDropdown(base);
        return {
            via: settled() ? 'native' : 'native (display did not update)',
            text: display() || wanted,
            closed: how
        };
    }

    // ---------- the fill routine ------------------------------------------
    async function fill(p, step) {
        // --- template form ---
        const tName = await waitFor('#templateName');
        const wrote = await typeInto(tName, p.name);
        step('Template Name', 'ok', wrote);
        await sleep(STEP_PAUSE);

        try {
            const r = await selectDropdown('orgAcct', () => true);   // first (only) option
            step('Originating Account', 'ok', `${r.text} · via ${r.via} · closed: ${r.closed}`);
        } catch (e) {
            step('Originating Account', 'warn', 'pick manually — ' + e.message);
        }
        await sleep(STEP_PAUSE * 2);

        // Company ID cascades from the originating account — only touch it if it didn't.
        const cid = () => (document.getElementById('achCompanyId_displayText')?.textContent || '').trim();
        if (cid() && !/^select an/i.test(cid())) {
            step('Originating ACH Company ID', 'ok', cid() + ' · auto-filled');
        } else {
            try {
                const r = await selectDropdown('achCompanyId', () => true);
                step('Originating ACH Company ID', 'ok', `${r.text} · via ${r.via} · closed: ${r.closed}`);
            } catch (e) {
                step('Originating ACH Company ID', 'warn', 'pick manually — ' + e.message);
            }
        }
        await sleep(STEP_PAUSE);

        const desc = await waitFor('#companyEntryDesc');
        step('Company Entry Description', 'ok', await typeInto(desc, ENTRY_DESC));
        await sleep(STEP_PAUSE);

        // --- recipient modal ---
        if (!AUTO_CREATE_NEW) {
            step('Create New', 'warn', 'AUTO_CREATE_NEW is off — click it yourself');
            return;
        }
        clickIt(await waitFor('#btnNewCreateRecipient'));
        step('Create New', 'ok', 'modal opening');

        const rName = await waitFor('#recipientName', { timeout: 15000 });
        step('Recipient Name', 'ok', await typeInto(rName, p.name));
        await sleep(STEP_PAUSE);

        const rId = await waitFor('#recipientId');
        step('Recipient ID', 'ok', await typeInto(rId, p.name));   // truncates at maxlength
        await sleep(STEP_PAUSE);

        try {
            const r = await selectDropdown('accType',
                txt => txt.toLowerCase() === p.type.toLowerCase());
            step('Account Type', 'ok', `${r.text} · via ${r.via} · closed: ${r.closed}`);
        } catch (e) {
            step('Account Type', 'warn', 'pick manually — ' + e.message);
        }
        await sleep(STEP_PAUSE);

        const acct = await waitFor('#accountNumber');
        step('Account Number', 'ok', await typeInto(acct, p.account));
        await sleep(STEP_PAUSE);

        // Bank: switch to "Enter Bank Information with Bank ID"
        const radio = await waitFor('#bankOption_enter_bankrecipient-radio-option-input');
        if (!radio.checked) clickIt(radio);
        step('Bank option', 'ok', 'enter bank ID');

        try {
            const bankId = await waitFor('#recipient_bankId', {
                timeout: 8000, test: visible
            });
            step('Bank ID', 'ok', await typeInto(bankId, p.routing));
        } catch (e) {
            step('Bank ID', 'warn', 'subform did not open — ' + e.message);
        }
        await sleep(STEP_PAUSE);

        const save = await waitFor('#input-saveToMaster-checkbox');
        if (!save.checked) clickIt(save);
        step('Save to Master Recipient List', save.checked ? 'ok' : 'warn',
             save.checked ? 'checked' : 'did not take');

        if (AUTO_CONTINUE) {
            clickIt(await waitFor('#btnContinueRecipient'));
            step('Continue', 'ok', 'submitted');
        } else {
            step('Continue', 'skip', 'review, then click it yourself');
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

    // ---------- panel ------------------------------------------------------
    const PANEL_ID = 'tt-ach-panel';

    function buildPanel() {
        if (document.getElementById(PANEL_ID)) return;

        const wrap = document.createElement('div');
        wrap.id = PANEL_ID;
        wrap.style.cssText = `
            position:fixed; right:16px; bottom:16px; width:330px; z-index:2147483000;
            background:#fff; border:1px solid #d4d4d8; border-radius:6px;
            box-shadow:0 4px 16px rgba(0,0,0,.18);
            font:13px/1.45 -apple-system,Segoe UI,Roboto,sans-serif; color:#18181b;`;

        const head = document.createElement('div');
        head.style.cssText = `padding:9px 12px; border-bottom:1px solid #e4e4e7;
            display:flex; justify-content:space-between; align-items:center; font-weight:600;`;
        head.textContent = 'Trade Ticket → ACH Template';

        const hide = document.createElement('button');
        hide.textContent = '–';
        hide.style.cssText = 'border:0;background:none;font-size:17px;cursor:pointer;line-height:1;';
        head.appendChild(hide);

        const body = document.createElement('div');
        body.style.cssText = 'padding:12px;';

        const ta = document.createElement('textarea');
        ta.placeholder = 'Paste the payload here…';
        ta.style.cssText = `width:100%; height:60px; padding:7px; resize:vertical;
            border:1px solid #d4d4d8; border-radius:4px;
            font:12px ui-monospace,Menlo,monospace;`;

        const go = document.createElement('button');
        go.textContent = 'Fill form';
        go.style.cssText = `width:100%; margin-top:8px; padding:9px; border:0; border-radius:5px;
            background:#18181b; color:#fff; font-size:13px; cursor:pointer;`;

        const logBox = document.createElement('div');
        logBox.style.cssText = 'margin-top:10px; max-height:260px; overflow:auto;';

        body.append(ta, go, logBox);
        wrap.append(head, body);
        document.body.appendChild(wrap);

        hide.addEventListener('click', () => {
            const shown = body.style.display !== 'none';
            body.style.display = shown ? 'none' : 'block';
            hide.textContent = shown ? '+' : '–';
        });

        const MARK = { ok: ['✓', '#16a34a'], warn: ['!', '#d97706'],
                       skip: ['–', '#a1a1aa'], err: ['✗', '#dc2626'] };

        const step = (label, state, detail) => {
            const [glyph, color] = MARK[state] || MARK.skip;
            const row = document.createElement('div');
            row.style.cssText = 'display:flex; gap:7px; margin-bottom:5px; font-size:12px;';
            row.innerHTML = `<span style="color:${color};font-weight:700">${glyph}</span>` +
                            `<span><b>${label}</b>` +
                            (detail ? `<div style="color:#71717a;font-size:11px">${detail}</div>` : '') +
                            `</span>`;
            logBox.appendChild(row);
            logBox.scrollTop = logBox.scrollHeight;
        };

        go.addEventListener('click', async () => {
            logBox.textContent = '';
            const p = parsePayload(ta.value);
            if (!p) {
                step('Payload', 'err', "couldn't parse — expected: OES # … Name RN ######### AN ##### Checking $0.00");
                return;
            }
            step('Payload', 'ok', `${p.name} · ${p.type} · ${p.account} · $${p.amount}`);
            go.disabled = true;
            go.textContent = 'Filling…';
            try {
                await fill(p, step);
            } catch (e) {
                step('Stopped', 'err', e.message);
            } finally {
                go.disabled = false;
                go.textContent = 'Fill form';
            }
        });
    }

    function removePanel() {
        document.getElementById(PANEL_ID)?.remove();
    }

    // Only surface on the ACH template create form (SPA — no page loads).
    const observer = new MutationObserver(() => {
        retargetSuccessButton();   // success screen can outlive the form
        if (document.querySelector('#templateName')) buildPanel();
        else removePanel();
    });
    observer.observe(document.body, { childList: true, subtree: true });
    if (document.querySelector('#templateName')) buildPanel();
})();