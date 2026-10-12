// ==UserScript==
// @name         RAD Row Counter
// @description  Adds # counter to items gone through RAD
// @namespace    tpa-tools
// @match        https://www.partners.example.com/*
// @run-at       document-idle
// @grant        none
// @version      1.3
// ==/UserScript==

(function () {
    'use strict';

    const TAG = '[TPA-Counter]';
    const COUNTER_ID = 'tpa-row-counter';

    // Only run inside the "main" frame (or top-level if it ever shows there).
    // This avoids running in topFramenew2 / slideMenu frames where it's pointless.
    const isMainFrame = window.name === 'main' || window.top === window.self;
    if (!isMainFrame) {
        // Quietly bail in irrelevant frames. Comment out the next line if you want to confirm.
        // console.log(`${TAG} Skipping frame:`, window.name, window.location.href);
        return;
    }

    console.log(`${TAG} Active in frame:`, window.name || '(top)', '| URL:', window.location.href);

    function tryInjectCounter(reason) {
        const table = document.getElementById('statusTable');
        if (!table) return false;

        const tbody = table.querySelector('tbody');
        if (!tbody) return false;

        const rows = Array.from(tbody.querySelectorAll('tr')).filter(tr => tr.querySelector('td'));
        if (rows.length === 0) return false;

        const successMsg = document.getElementById('successMsg');
        if (!successMsg) return false;

        const p = successMsg.querySelector('p');
        if (!p) return false;

        let counter = document.getElementById(COUNTER_ID);
        if (!counter) {
            counter = document.createElement('span');
            counter.id = COUNTER_ID;
            counter.style.cssText = 'margin-left:10px; padding:2px 8px; background:#28a745; color:#fff; border-radius:10px; font-weight:bold; font-size:13px;';
            p.appendChild(counter);
        }
        counter.textContent = `Total: ${rows.length}`;
        console.log(`${TAG} ✅ Injected (${reason}). Total: ${rows.length}`);
        return true;
    }

    // Initial attempt
    if (tryInjectCounter('initial')) return;

    // The "main" frame navigates to different pages — observer will see the TPA
    // page's content get added when the user navigates to it and runs the process.
    console.log(`${TAG} Setting up observer in main frame.`);
    const observer = new MutationObserver(() => {
        tryInjectCounter('mutation');
        // Don't disconnect — main frame can navigate away and back, or the user
        // can re-run the TPA process. Leave observer running for the session.
    });

    observer.observe(document.body, {
        childList: true,
        subtree: true
    });
})();