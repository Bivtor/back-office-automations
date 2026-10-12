// ==UserScript==
// @name         OES Internal Note — Quick Fill
// @namespace    http://tampermonkey.net/
// @version      1.0
// @description  Adds one-click canned notes to the Add Internal Note modal
// @match        https://oes.example.com/
// @match        https://oes.example.com/index.cfm*
// @grant        none
// @run-at       document-idle
// ==/UserScript==

(function () {
    'use strict';

    // ===== CONFIG ========================================================
    // {OES} is replaced with the OES ID from the form's hidden field.
    const NOTES = [
        'Funds requested in UMB.',
        'Void Check in OES #, Template requested in UMB.',
        'Template requested in UMB, using TT',
    ];

    // Short button labels, in the same order as NOTES.
    const LABELS = ['Funds requested', 'Void + template', 'Template w/ TT'];

    const ONLY_INTERNAL = true;   // false -> also show on NIGO notes
    const NAV_STATE     = 56;     // Purchase Team -> changeState(56)
    // =====================================================================

    const WRAP_ID = 'pm-note-quickfill';

    function navIsActive() {
        const active = document.querySelector('#nav .navBox.activeNav');
        if (!active) return false;
        const re = new RegExp('changeState\\(\\s*' + NAV_STATE + '\\s*\\)');
        return re.test(active.getAttribute('onclick') || '');
    }

    function fillNote(form, text) {
        const ta = form.querySelector('textarea[name="notes"]');
        if (!ta) return;

        const oes = form.querySelector('input[name="OESID"]')?.value || '';
        ta.value = text.replace(/\{OES\}/g, oes);

        ta.dispatchEvent(new Event('input',  { bubbles: true }));
        ta.dispatchEvent(new Event('change', { bubbles: true }));
        ta.focus();
        ta.setSelectionRange(ta.value.length, ta.value.length);
    }

    function inject(form) {
        const existing = form.querySelector('#' + WRAP_ID);

        if (!navIsActive()) {
            existing?.remove();      // left Purchase Team with the modal still around
            return;
        }
        if (existing) return;

        if (ONLY_INTERNAL) {
            const type = form.querySelector('input[name="noteType"]')?.value || '';
            if (type !== 'INTERNAL_NOTE') return;
        }

        const select = form.querySelector('select[name="docID"]');
        if (!select) return;
        const anchor = select.closest('div');
        if (!anchor) return;

        const wrap = document.createElement('div');
        wrap.id = WRAP_ID;
        wrap.style.cssText = 'margin:10px 5px 0 5px; display:flex; flex-wrap:wrap; gap:6px;';

        NOTES.forEach((text, i) => {
            const b = document.createElement('button');
            b.type = 'button';
            b.className = 'btn btn-outline-secondary btn-sm';
            b.textContent = LABELS[i] || `Note ${i + 1}`;
            b.title = text;
            b.addEventListener('click', (e) => {
                e.preventDefault();
                b.blur();
                fillNote(form, text);
            });
            wrap.appendChild(b);
        });

        anchor.insertAdjacentElement('afterend', wrap);
    }

    const observer = new MutationObserver(() => {
        document.querySelectorAll('form[name="noteForm"]').forEach(inject);
    });
    observer.observe(document.body, { childList: true, subtree: true });

    document.querySelectorAll('form[name="noteForm"]').forEach(inject);
})();