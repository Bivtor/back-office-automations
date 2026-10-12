// ==UserScript==
// @name         OES Alt+Click Date Fill
// @namespace    http://tampermonkey.net/
// @version      1.0
// @description  Alt+click any input to insert today's date (m/d/yyyy)
// @author       You
// @include      https://www.partners.example.com/SimpleAlloc/Maintenance/rolloverEdit.cfm*
// @grant        none
// ==/UserScript==

(function () {
    'use strict';

    function todayFormatted() {
        const d = new Date();
        return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`;
    }

    document.addEventListener('click', (e) => {
        if (!e.altKey) return;
        const el = e.target;
        if (el.tagName !== 'INPUT') return;

        e.preventDefault();
        el.value = todayFormatted();
        el.dispatchEvent(new Event('input',  { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        el.dispatchEvent(new FocusEvent('blur', { bubbles: true }));
    }, true);

})();