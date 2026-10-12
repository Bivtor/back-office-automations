// ==UserScript==
// @name         OES Better Documents List
// @note         The upstream app has since shipped its own sort; kept as the original workaround.
// @namespace    http://tampermonkey.net/
// @version      2.0
// @description  Alphabetizes every "Current Documents" list on the OES page — main viewer, fax modal, and any other instance — by anchoring on the td.currentDocs header that all of them share. Case-insensitive natural-numeric sort. Adds an asc/desc caret. SPA-aware via MutationObserver.
// @match        https://oes.example.com/
// @match        https://oes.example.com/index.cfm?openModule=37
// @grant        none
// ==/UserScript==

(function () {
    'use strict';

    const PROCESSED_FLAG = 'data-oes-sorted';
    const CARET_CLASS = 'pm-doc-sort-caret';

    // ---- Comparator -------------------------------------------------------
    // Case-insensitive + natural numeric order:
    //   - groups "0-", "1-", "2-" ... prefixed docs at the top in number order
    //   - sorts "Rollover 1/2/10" naturally (not 1,10,2)
    //   - puts lowercase-initial names (eSign, onboarding) in their correct A-Z spot
    const collator = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' });

    function sortKey(text) {
        return text; // default: sort on the full display text - there was previously another line where the AI thought it might be a good idea to add a potential sort comparator that strips "#-" prefixes
    }

    function compareText(a, b) {
        return collator.compare(sortKey(a), sortKey(b));
    }

    // ---- Generic row helpers ---------------------------------------------
    function isDocRow(row) {
        return row.classList.contains('d0') || row.classList.contains('d1');
    }

    // Name cell differs per table (anchor in the viewer, plain text in the modal).
    // The name is always the longest-text cell in the row, so just take that.
    function getDisplayText(row) {
        let best = '';
        for (const td of row.children) {
            const t = td.textContent.trim();
            if (t.length > best.length) best = t;
        }
        return best;
    }

    function sortDocTable(table, ascending) {
        const tbody = table.querySelector('tbody');
        if (!tbody) return;

        const docRows = Array.from(tbody.children).filter(isDocRow);
        if (docRows.length < 2) return;

        // The doc rows are a contiguous block. Capture the node right after the
        // block (spacer row, or null) so we can reinsert without disturbing the
        // header, the OES-ID "title" row, or the spacer.
        const after = docRows[docRows.length - 1].nextSibling;

        const sorted = docRows.slice().sort((a, b) => {
            const cmp = compareText(getDisplayText(a), getDisplayText(b));
            return ascending ? cmp : -cmp;
        });

        docRows.forEach(r => r.remove());
        for (const row of sorted) tbody.insertBefore(row, after); // null -> append

        // Re-apply alternating zebra classes.
        sorted.forEach((row, i) => {
            row.classList.remove('d0', 'd1');
            row.classList.add(i % 2 === 0 ? 'd0' : 'd1');
        });
    }

    function makeCaret(ascending) {
        const span = document.createElement('span');
        span.className = CARET_CLASS;
        span.style.cssText = 'cursor:pointer;margin-left:8px;user-select:none;font-size:11px;display:inline-block;';
        span.title = 'Toggle sort order (A-Z / Z-A)';
        span.textContent = ascending ? '▲' : '▼';
        return span;
    }

    function processTable(table) {
        if (table.getAttribute(PROCESSED_FLAG) === '1') return;

        // Signature shared by every instance of this list.
        const headerCell = table.querySelector('td.currentDocs');
        if (!headerCell) return;
        if (table.querySelectorAll('tr.d0, tr.d1').length < 2) return;

        table.setAttribute(PROCESSED_FLAG, '1');

        let ascending = true;

        if (!headerCell.querySelector('.' + CARET_CLASS)) {
            const caret = makeCaret(ascending);
            caret.addEventListener('click', (e) => {
                e.stopPropagation();
                ascending = !ascending;
                caret.textContent = ascending ? '▲' : '▼';
                sortDocTable(table, ascending);
            });
            headerCell.appendChild(caret);
        }

        sortDocTable(table, ascending); // initial ascending sort
    }

    function scan() {
        // td.currentDocs is the universal anchor — finds the viewer table, the fax
        // modal table, and any third instance, wherever they're wrapped.
        document.querySelectorAll('td.currentDocs').forEach(cell => {
            const table = cell.closest('table');
            if (table) processTable(table);
        });
    }

    // Coalesce bursts of ExtJS mutations into one scan per frame.
    let scheduled = false;
    const observer = new MutationObserver(() => {
        if (scheduled) return;
        scheduled = true;
        requestAnimationFrame(() => { scheduled = false; scan(); });
    });
    observer.observe(document.body, { childList: true, subtree: true });

    scan();
})();