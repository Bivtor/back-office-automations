// ==UserScript==
// @name         SimpleAllo Paste Splitter
// @namespace    http://tampermonkey.net/
// @version      1.1
// @description  Paste Excel column of dollar amounts into split deposit inputs
// @author       You
// @include      https://www.partners.example.com/SimpleAlloc/DailyDeposit/DepositSplitACH.cfm*
// @grant        none
// ==/UserScript==

(function () {
    'use strict';

    function cleanValue(raw) {
        return raw.replace(/[$,\s]/g, '').trim();
    }

    function parseClipboard(text) {
        return text
            .split('\n')
            .map(cleanValue)
            .filter(v => v.length > 0 && !isNaN(v) && v !== '');
    }

    function triggerEvents(el) {
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        // Also trigger blur so calcRemaining() fires on each field
        el.dispatchEvent(new FocusEvent('blur', { bubbles: true }));
    }

    function getAddButton() {
        return document.getElementById('AddSplitACHBtn');
    }

    function getInput(n) {
        return document.getElementById('SplitACH_Amount_' + n);
    }

    function getCurrentCount() {
        const counter = document.getElementById('SplitACHCount');
        return counter ? parseInt(counter.value, 10) : 1;
    }

    function clickAdd() {
        const btn = getAddButton();
        if (btn) btn.click();
    }

    function fillValues(values) {
        // Fill row 1 which always exists
        const first = getInput(1);
        if (!first) return;

        // How many rows exist right now
        let count = getCurrentCount();

        for (let i = 0; i < values.length; i++) {
            const rowNum = i + 1;

            // Add a new row if needed
            if (rowNum > count) {
                clickAdd();
                count = getCurrentCount(); // button click updates #SplitACHCount synchronously
            }

            const input = getInput(rowNum);
            if (!input) {
                console.warn('[SplitPaste] Could not find input for row', rowNum);
                break;
            }

            input.value = values[i];
            triggerEvents(input);
        }

        console.log('[SplitPaste] Filled', values.length, 'value(s).');
    }

    function onPaste(e) {
        // Only intercept paste on one of our split amount inputs
        const target = e.target;
        if (!target || !target.id || !target.id.startsWith('SplitACH_Amount_')) return;

        const text = (e.clipboardData || window.clipboardData).getData('text');
        const values = parseClipboard(text);

        // Single value — let it paste normally
        if (values.length < 2) return;

        e.preventDefault();
        fillValues(values);
    }

    window.addEventListener('load', () => {
        document.addEventListener('paste', onPaste, true);
        console.log('[SplitPaste] Ready.');
    });

})();