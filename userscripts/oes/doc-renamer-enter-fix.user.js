// ==UserScript==
// @name         OES Doc Renamer "Enter Bug" fix
// @note         The upstream app has since fixed this; kept as the original workaround.
// @namespace    http://tampermonkey.net/
// @version      1.0
// @description  Makes Enter key save the document rename instead of navigating away
// @match        https://oes.example.com/
// @match        https://oes.example.com/index.cfm?openModule=37
// @grant        none
// ==/UserScript==

(function () {
    'use strict';

    function handleEnter(e) {
        if (e.key !== 'Enter') return;
        e.preventDefault();
        e.stopPropagation();
        e.target.blur(); // triggers their save handler
    }

    const observer = new MutationObserver((mutations) => {
        for (const mutation of mutations) {
            // Input appeared — we just entered edit mode
            for (const node of mutation.addedNodes) {
                if (node.nodeType !== 1) continue;
                const inputs = node.tagName === 'INPUT' ? [node] : Array.from(node.querySelectorAll('input'));
                for (const input of inputs) {
                    if (!input.closest('.docTable')) continue; // only doc table inputs
                    input.addEventListener('keydown', handleEnter);
                }
            }

            // Input removed — edit mode exited, nothing to clean up
            // (listener dies with the element)
        }
    });

    observer.observe(document.body, { childList: true, subtree: true });

})();