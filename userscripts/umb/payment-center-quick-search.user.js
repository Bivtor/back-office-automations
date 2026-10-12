// ==UserScript==
// @name         UMB Payment Center Quick Search
// @namespace    http://tampermonkey.net/
// @version      1.0
// @description  Adds one-click search presets (Today's TT's + Loan Pulls / Today's ACHs) to the UMB Payment Center
// @match        https://umbdirect.umb.com/rwd-web/*
// @grant        none
// ==/UserScript==

(function () {
    'use strict';

    const WRAP_ID = 'pm-quick-search';
    const P = 'alltabSearch'; // id prefix for the ALL tab's search form
    const sleep = ms => new Promise(r => setTimeout(r, ms));

    // ---------- Dates / business days ----------

    const pad = n => String(n).padStart(2, '0');
    const fmt = d => `${pad(d.getMonth() + 1)}/${pad(d.getDate())}/${d.getFullYear()}`;
    const dayKey = d => `${d.getFullYear()}-${d.getMonth()}-${d.getDate()}`;

    function nthWeekday(y, m, wd, n) {
        const d = new Date(y, m, 1);
        d.setDate(1 + ((wd - d.getDay() + 7) % 7) + (n - 1) * 7);
        return d;
    }

    function lastWeekday(y, m, wd) {
        const d = new Date(y, m + 1, 0);
        d.setDate(d.getDate() - ((d.getDay() - wd + 7) % 7));
        return d;
    }

    // Federal Reserve holidays (no ACH settlement)
    function fedHolidays(y) {
        const fixed = [[0, 1], [5, 19], [6, 4], [10, 11], [11, 25]].map(([m, dd]) => {
            const d = new Date(y, m, dd);
            if (d.getDay() === 0) d.setDate(d.getDate() + 1); // Sunday -> observed Monday
            return d; // Saturday: Fed stays open Friday, and Saturday is already skipped
        });
        const floating = [
            nthWeekday(y, 0, 1, 3),  // MLK Day
            nthWeekday(y, 1, 1, 3),  // Presidents Day
            lastWeekday(y, 4, 1),    // Memorial Day
            nthWeekday(y, 8, 1, 1),  // Labor Day
            nthWeekday(y, 9, 1, 2),  // Columbus Day
            nthWeekday(y, 10, 4, 4), // Thanksgiving
        ];
        return new Set([...fixed, ...floating].map(dayKey));
    }

    function isBusinessDay(d) {
        const wd = d.getDay();
        return wd !== 0 && wd !== 6 && !fedHolidays(d.getFullYear()).has(dayKey(d));
    }

    function nextBusinessDay(from) {
        const d = new Date(from.getFullYear(), from.getMonth(), from.getDate());
        do { d.setDate(d.getDate() + 1); } while (!isBusinessDay(d));
        return d;
    }

    // ---------- Form helpers ----------

    function setInput(id, value) {
        const el = document.getElementById(id);
        if (!el) throw new Error(`#${id}`);
        const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
        el.focus();
        setter.call(el, value); // native setter so Angular sees the change
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        el.blur();
    }

    const displayText = base =>
        document.getElementById(`${base}_displayText`)?.textContent.trim() ?? '';

    // Option labels live on the hidden accessible listbox; li id = base + 'hiddenOption' + value
    function optionValueFor(base, re) {
        const li = [...document.querySelectorAll(`#${base}hiddenListbox li[role="option"]`)]
            .find(li => re.test((li.getAttribute('aria-label') || li.textContent).trim()));
        return li ? li.id.slice(`${base}hiddenOption`.length) : null;
    }

    function visibleOption(re) {
        return [...document.querySelectorAll('[role="option"]')]
            .filter(el => !el.closest('.dropdown-sr-only-items') && el.offsetParent !== null)
            .find(el => re.test(el.textContent.trim()));
    }

    async function setDropdown(base, re, label) {
        if (re.test(displayText(base))) return; // already set

        const select = document.getElementById(`${base}select`);
        const opener = document.querySelector(`#${base} a.selected-value`);
        let opened = false;
        let value = optionValueFor(base, re);

        // Options can be lazy-rendered until the dropdown is opened
        if (value === null && opener) {
            opener.click();
            opened = true;
            await sleep(300);
            value = optionValueFor(base, re);
        }

        // Primary: set the native select Angular is bound to
        if (value !== null && select) {
            select.value = value;
            select.dispatchEvent(new Event('change', { bubbles: true }));
            await sleep(100);
        }

        // Fallback: click the option in the visible custom dropdown
        if (!re.test(displayText(base))) {
            let opt = visibleOption(re);
            if (!opt && opener && !opened) {
                opener.click();
                await sleep(300);
                opt = visibleOption(re);
            }
            opt?.click();
            await sleep(100);
        } else if (opened) {
            document.body.click(); // close the dropdown we opened
        }

        if (!re.test(displayText(base))) throw new Error(label);
    }

    // ---------- Presets ----------

    async function runPreset(getDate, btn) {
        const ds = fmt(getDate());
        btn.disabled = true;
        try {
            await setDropdown(`${P}searchDateType`, /^payment date$/i, 'Date Type');
            setInput(`${P}searchFrom-datepicker`, ds);
            setInput(`${P}searchTo-datepicker`, ds);
            await setDropdown(`${P}searchPaymentType`, /all\s+ach/i, 'Payment Type');
            await sleep(150);
            const search = document.getElementById(`${P}btnSearch`);
            if (!search) throw new Error('Search button');
            search.click();
        } catch (e) {
            alert(`Quick search failed — could not set: ${e.message}`);
        } finally {
            btn.disabled = false;
        }
    }

    const PRESETS = [
        { label: "Today's TT's + Loan Pulls", getDate: () => nextBusinessDay(new Date()) },
        { label: "Today's ACHs",              getDate: () => new Date() },
    ];

    // ---------- Injection ----------

    function inject() {
        if (document.getElementById(WRAP_ID)) return;
        const form = document.getElementById(`${P}btnSearch`)?.closest('form');
        if (!form) return;

        const wrap = document.createElement('div');
        wrap.id = WRAP_ID;
        wrap.className = 'col-xs-12';
        wrap.style.cssText = 'clear:both; display:flex; gap:10px; margin:6px 0 12px; padding-left:15px;';

        PRESETS.forEach(({ label, getDate }) => {
            const btn = document.createElement('button');
            btn.type = 'button'; // don't submit the form directly
            btn.className = 'btn btn-outline';
            btn.textContent = label;
            btn.addEventListener('mouseenter', () => { btn.title = `Payment date: ${fmt(getDate())}`; });
            btn.addEventListener('click', () => runPreset(getDate, btn));
            wrap.appendChild(btn);
        });

        form.appendChild(wrap);
    }

    let scheduled = false;
    new MutationObserver(() => {
        if (scheduled) return;
        scheduled = true;
        requestAnimationFrame(() => { scheduled = false; inject(); });
    }).observe(document.body, { childList: true, subtree: true });

    inject();
})();