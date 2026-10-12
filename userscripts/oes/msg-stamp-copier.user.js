// ==UserScript==
// @name         OES MSG Stamp Copier
// @namespace    http://tampermonkey.net/
// @version      1.1
// @description  Injects a clipboard copy button into the Documents card header
// @match https://oes.example.com/
// @match https://oes.example.com/index.cfm?openModule=37
// @grant        GM_setClipboard
// ==/UserScript==

(function () {
    'use strict';

    const BUTTON_ID = 'pm-copy-row-btn';

    function formatDate(d) {
        return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`;
    }

    function extractData() {
        const errors = [];

        // Col 0: Today's date
        const col0 = formatDate(new Date());

        // Col 1: Client name
        const clientDiv = [...document.querySelectorAll('#editHeader .col-sm-4.text-start')]
        .find(el => el.querySelector('b')?.textContent.trim() === 'Client:');
        const col1 = clientDiv?.querySelector('a.clickable')?.innerText.trim() ?? (errors.push('Client'), '');

        // Col 2: Hardcoded
        const col2 = 'Retirement Transfer';

        // Col 3 & 5 — target the specific rollover table
        const firstRolloverRow = document.querySelector('table.roInfoTable tbody tr:not(.title)');
        const expTd = firstRolloverRow?.querySelectorAll('td')[7];
        const col3 = expTd ? expTd.innerText.trim().replace('$', '') : (errors.push('Exp $'), '');


        // Col 4: Rep name — second row's col-sm-5 in #editHeader (first row is "Current State")
        const repDiv = [...document.querySelectorAll('#editHeader .row')]
        .find(row => [...row.querySelectorAll('strong')].some(s => s.textContent.trim() === 'Rep:'));
        const repSpan = [...(repDiv?.querySelectorAll('span') ?? [])]
        .find(s => !s.className.includes('fa'));
        const col4 = repSpan?.innerText.trim() ?? (errors.push('Rep name'), '');

        // Col 5
        const col5 = firstRolloverRow?.querySelectorAll('td')[1]?.innerText.trim() ?? (errors.push('Company'), '');

        // Col 6: OES ID — first col-sm-4.text-start in #editHeader, strip label
        const oesDiv = document.querySelector('#editHeader .col-sm-4.text-start');
        const col6 = oesDiv
        ? oesDiv.innerText.replace('OES ID:', '').trim()
        : (errors.push('OES ID'), '');

        if (errors.length > 0) {
            alert(`Copy failed — could not find: ${errors.join(', ')}`);
            return null;
        }

        return [col0, col1, col2, col3, col4, col5, col6].join('\t');
    }

    function injectButton(headerDiv) {
        if (document.getElementById(BUTTON_ID)) return;

        const wrapper = document.createElement('span');
        wrapper.id = BUTTON_ID;
        wrapper.title = 'Copy row to clipboard';
        wrapper.style.cssText = 'cursor:pointer; margin-top:2px; margin-left:5px; color:white; display:inline-flex; align-items:center;';
        wrapper.innerHTML = `
            <svg width="18" height="18" viewBox="40 40 175 175" xmlns="http://www.w3.org/2000/svg">
                <path d="M48.186 92.137c0-8.392 6.49-14.89 16.264-14.89s29.827-.225 29.827-.225-.306-6.99-.306-15.88c0-8.888 7.954-14.96 17.49-14.96 9.538 0 56.786.401 61.422.401 4.636 0 8.397 1.719 13.594 5.67 5.196 3.953 13.052 10.56 16.942 14.962 3.89 4.402 5.532 6.972 5.532 10.604 0 3.633 0 76.856-.06 85.34-.059 8.485-7.877 14.757-17.134 14.881-9.257.124-29.135.124-29.135.124s.466 6.275.466 15.15-8.106 15.811-17.317 16.056c-9.21.245-71.944-.49-80.884-.245-8.94.245-16.975-6.794-16.975-15.422s.274-93.175.274-101.566zm16.734 3.946l-1.152 92.853a3.96 3.96 0 0 0 3.958 4.012l73.913.22a3.865 3.865 0 0 0 3.91-3.978l-.218-8.892a1.988 1.988 0 0 0-2.046-1.953s-21.866.64-31.767.293c-9.902-.348-16.672-6.807-16.675-15.516-.003-8.709.003-69.142.003-69.142a1.989 1.989 0 0 0-2.007-1.993l-23.871.082a4.077 4.077 0 0 0-4.048 4.014zm106.508-35.258c-1.666-1.45-3.016-.84-3.016 1.372v17.255c0 1.106.894 2.007 1.997 2.013l20.868.101c2.204.011 2.641-1.156.976-2.606l-20.825-18.135zm-57.606.847a2.002 2.002 0 0 0-2.02 1.988l-.626 96.291a2.968 2.968 0 0 0 2.978 2.997l75.2-.186a2.054 2.054 0 0 0 2.044-2.012l1.268-62.421a1.951 1.951 0 0 0-1.96-2.004s-26.172.042-30.783.042c-4.611 0-7.535-2.222-7.535-6.482S152.3 63.92 152.3 63.92a2.033 2.033 0 0 0-2.015-2.018l-36.464-.23z" stroke="white" fill="white" fill-rule="evenodd"/>
            </svg>
        `;

        wrapper.addEventListener('click', () => {
            const tsv = extractData();
            if (!tsv) return;
            GM_setClipboard(tsv, 'text');

            wrapper.style.opacity = '0.4';
            wrapper.title = 'Copied!';
            setTimeout(() => {
                wrapper.style.opacity = '1';
                wrapper.title = 'Copy row to clipboard';
            }, 1200);
        });

        const notesHeader = document.querySelector('div.card-header.bg-secondary');
        if (notesHeader) {
            const notesTitle = notesHeader.querySelector('h4.editTitle');
            notesTitle.insertAdjacentElement('beforebegin', wrapper);
        }
    }

    function removeButton() {
        document.getElementById(BUTTON_ID)?.remove();
    }

    const observer = new MutationObserver(() => {
        const header = document.querySelector('div.card-header.bg-secondary');
        if (header) {
            injectButton(header);
        } else {
            removeButton();
        }
    });

    observer.observe(document.body, { childList: true, subtree: true });

    // Handle case where content already exists on load
    const header = document.querySelector('div.card-header.bg-secondary');
    if (header) injectButton(header);

})();