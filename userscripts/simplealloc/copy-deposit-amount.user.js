// ==UserScript==
// @name         SimpleAllo Copy Deposit Amount
// @namespace    http://tampermonkey.net/
// @version      1.0
// @description  Click total amount to copy clean numeric value to clipboard
// @author       You
// @include      https://www.partners.example.com/SimpleAlloc/DailyDeposit/DepositView.cfm*
// @grant        none
// ==/UserScript==

(function () {
    'use strict';

    window.addEventListener('load', () => {
        // Find the "Total Amount: $XX,XXX.XX" element in the panel heading
        const headingCols = document.querySelectorAll('.panel-heading .col-md-6');
        let amountEl = null;

        for (const col of headingCols) {
            if (col.textContent.includes('Total Amount:')) {
                amountEl = col;
                break;
            }
        }

        if (!amountEl) return;

        amountEl.style.cursor = 'pointer';

        amountEl.addEventListener('click', () => {
            const raw = amountEl.textContent;
            const clean = raw.replace(/[^0-9.]/g, '');
            navigator.clipboard.writeText(clean);
        });
    });

})();