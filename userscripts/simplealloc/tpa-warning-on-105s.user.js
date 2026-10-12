// ==UserScript==
// @name         OES Rollover - TPA Warning on 105s
// @namespace    http://tampermonkey.net/
// @version      1.0
// @description  Highlights Source Acct Type red when TPA + Transfer Type + Acct Type combo is flagged
// @match        https://www.partners.example.com/SimpleAlloc/Maintenance/rolloverEdit.cfm*
// @grant        none
// ==/UserScript==

(function () {
    'use strict';

    const FLAGGED_TPAS = [
        'TCG ADMINISTRATORS',
        'NATIONAL BENEFIT SERVICES (NBS)',
        'SCHOOLSFIRST FEDERAL CREDIT UNION'
    ];

    const tpaSelect = document.getElementById('fldTPAName');
    const transferType = document.getElementById('fldTranCodeID');
    const MSG_ID = 'oes-tpa-warning';

    function evaluate() {
        const tpa = tpaSelect?.value.trim().toUpperCase();
        const transfer = transferType?.options[transferType.selectedIndex]?.text.trim();
        const accountsTable = document.querySelectorAll('table.table')[2];

        const acct = accountsTable?.rows[2]?.cells[1]?.innerText.trim() ?? '';

        // console.log('TPA Validator | TPA:', tpa, '| Transfer:', transfer, '| Acct:', acct);

        const tpaFlagged = FLAGGED_TPAS.includes(tpa);
        const transferFlagged = transfer === '105 - ROLLOVER';
        const acctFlagged = acct.startsWith('403');
        const flagged = tpaFlagged && transferFlagged && acctFlagged;

       // console.log('TPA Validator | Flags:', { tpaFlagged, transferFlagged, acctFlagged, flagged });

        document.getElementById(MSG_ID)?.remove();

        if (flagged) {
            const accountsTable = [...document.querySelectorAll('table.table')]
            .find(t => t.innerText.includes('This rollover is split between the following accounts'));

            const targetTable = document.querySelectorAll('table.table')[2];

            if (targetTable) {
              //  console.log('TPA Validator | Injecting warning message');
                const row = document.createElement('tr');
                row.id = MSG_ID;
                const cell = document.createElement('td');
                cell.colSpan = 8;
                cell.textContent = '⚠ TPA not required for 105 into plan';
                cell.style.cssText = 'font-size:9px; color:red; font-style:italic; padding:3px 6px;';
                row.appendChild(cell);
                targetTable.querySelector('tbody').appendChild(row);
            } else {
                console.log('TPA Validator | Flagged but accounts table not found — message not injected');
            }
        } else {
          //  console.log('TPA Validator | Conditions not met — no message');
        }
    }

    tpaSelect?.addEventListener('change', evaluate);
    transferType?.addEventListener('change', evaluate);

    // Run on load in case values are pre-populated
    evaluate();

})();