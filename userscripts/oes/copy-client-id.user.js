// ==UserScript==
// @name         OES Copy Client ID
// @namespace    http://tampermonkey.net/
// @version      1.1
// @description  Click (or Alt+hover) the Client ID span to copy the full SSN from its title attribute
// @author       You
// @match        https://oes.example.com/
// @match        https://oes.example.com/index.cfm?openModule=37
// @grant        GM_setClipboard
// ==/UserScript==
(function () {
    'use strict';

    function copy(text) {
        // GM_setClipboard works without a click or window focus; the rest is just a safety net
        if (typeof GM_setClipboard === 'function') {
            GM_setClipboard(text, { type: 'text', mimetype: 'text/plain' });
        } else {
            navigator.clipboard.writeText(text);
        }
    }

    function valueFrom(span) {
        const title = span.getAttribute('title') || '';
        // "123-45-6789; PlayerID=1234567" -> part before the ";"
        return title.split(';')[0].trim();
    }

    function bindSpan(span) {
        if (span.dataset.ssnCopyBound) return; // don't double-bind on re-render
        span.dataset.ssnCopyBound = '1';

        span.addEventListener('click', () => {
            const value = valueFrom(span);
            if (value) copy(value);
        });

        // Alt + hover: copy without clicking, so you keep your always-on-top tab in place
        span.addEventListener('mousemove', (e) => {
            if (!e.altKey) { span.dataset.ssnHoverCopied = ''; return; }
            if (span.dataset.ssnHoverCopied === '1') return; // only copy once per hover
            const value = valueFrom(span);
            if (!value) return;
            copy(value);
            span.dataset.ssnHoverCopied = '1';
        });
        span.addEventListener('mouseleave', () => {
            span.dataset.ssnHoverCopied = '';
        });
    }

    function scan(root) {
        root.querySelectorAll('span[title*="PlayerID="]').forEach(bindSpan);
    }

    scan(document);

    const observer = new MutationObserver((mutations) => {
        for (const m of mutations) {
            for (const node of m.addedNodes) {
                if (node.nodeType !== 1) continue;
                if (node.matches && node.matches('span[title*="PlayerID="]')) bindSpan(node);
                if (node.querySelectorAll) scan(node);
            }
        }
    });
    observer.observe(document.body, { childList: true, subtree: true });
})();