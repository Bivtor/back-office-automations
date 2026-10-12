// ==UserScript==
// @name         OES TPA Form Flag
// @namespace    http://tampermonkey.net/
// @version      1.0
// @description  Flags TPAs that require a TPA form in the Employer Summary modal
// @match        https://oes.example.com/
// @match        https://oes.example.com/index.cfm?openModule=37
// @grant        none
// ==/UserScript==

(function () {
    'use strict';

    const FLAG_CLASS = 'pm-tpa-form-flag';
    // Case-insensitive "contains" match; \s* also catches "Schools First"
    const TPA_PATTERN = /schools\s*first|tcg|nbs/i;

    function flagTpa() {
        // TPA name lives in the first <div><b>...</b></div> of the TPA Information tab
        document.querySelectorAll('#empTab2 > div > b').forEach(nameEl => {
            const container = nameEl.parentElement;
            const existing = container.querySelector('.' + FLAG_CLASS);
            const matches = TPA_PATTERN.test(nameEl.textContent);

            if (matches && !existing) {
                const flag = document.createElement('span');
                flag.className = FLAG_CLASS;
                flag.textContent = ' - TPA Form Required';
                flag.style.cssText = 'color:red; font-weight:bold;';
                nameEl.insertAdjacentElement('afterend', flag);
            } else if (!matches && existing) {
                // Modal content was swapped to a different employer
                existing.remove();
            }
        });
    }

    // Only does work when the modal's TPA tab is actually in the DOM;
    // batched to one check per frame so it stays cheap.
    let scheduled = false;
    const observer = new MutationObserver(() => {
        if (scheduled) return;
        scheduled = true;
        requestAnimationFrame(() => {
            scheduled = false;
            flagTpa();
        });
    });

    observer.observe(document.body, { childList: true, subtree: true });

    // Handle case where the modal is already open on load
    flagTpa();
})();