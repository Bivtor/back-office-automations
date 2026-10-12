// ==UserScript==
// @name         OES Better Uploader
// @namespace    http://tampermonkey.net/
// @version      1.6
// @description  Auto-fills display name, smart doc-type defaults (rtf-msg -> Retirement Transfer, equitable -> Equitable Transfer Request, else Unknown), auto-opens file picker, Enter to upload, auto-refresh on success
// @match        https://oes.example.com/
// @match        https://oes.example.com/index.cfm?openModule=37
// @grant        none
// ==/UserScript==
(function () {
    'use strict';
    let pendingFilePick = false;

    // Doc-type rules: normalized name (lowercased, spaces stripped) -> {id, label}
    function pickDocType(rawName) {
        const n = (rawName || '').toLowerCase().replace(/\s+/g, '');
        if (n.includes('rtf-msg'))    return { id: '240', label: 'Retirement Transfer' };
        if (n.includes('equitable'))  return { id: '408', label: 'Equitable Transfer Request' };
        return { id: '481', label: 'Unknown Document' };
    }
    function applyDocType(rawName, docSearch, selectedDoc) {
        const t = pickDocType(rawName);
        selectedDoc.value = t.id;
        docSearch.value = t.label;
    }

    function hookLoadDoc() {
        const win = (typeof unsafeWindow !== 'undefined') ? unsafeWindow : window;
        if (typeof win.loadDoc === 'undefined') return;
        const original = win.loadDoc;
        win.loadDoc = function(...args) {
            pendingFilePick = true;
            return original.apply(this, args);
        };
    }
    function hookUploadBtn() {
        const uploadBtn = document.querySelector('#uploadFooter .btn-success');
        if (!uploadBtn || uploadBtn._uploadHooked) return;
        uploadBtn._uploadHooked = true;
        uploadBtn.addEventListener('click', () => {
            const oesIDField = document.querySelector('#uploadDiv input[name="OESID"]');
            const oesID = oesIDField ? oesIDField.value : null;
            if (!oesID) return;
            setTimeout(() => {
                if (typeof window.closeManagement === 'function') {
                    window.closeManagement(oesID);
                }
            }, 3000);
        });
    }
    function isModalOpen() {
        const uploadDiv = document.getElementById('uploadDiv');
        if (!uploadDiv) return false;
        return uploadDiv.offsetParent !== null && uploadDiv.style.display !== 'none';
    }
    function handleEnter(e) {
        if (e.key !== 'Enter') return;
        if (!isModalOpen()) return;
        const focused = document.activeElement;
        if (!focused || !document.getElementById('uploadDiv').contains(focused)) return;
        const fileInput = document.getElementById('selectedFile');
        if (!fileInput || !fileInput.files || fileInput.files.length === 0) return;
        e.preventDefault();
        e.stopPropagation();
        const uploadBtn = document.querySelector('#uploadFooter .btn-success');
        if (uploadBtn) uploadBtn.click();
    }
    function enhanceModal() {
        const fileInput = document.getElementById('selectedFile');
        const displayInput = document.querySelector('input[name="displayName"]');
        const docSearch = document.getElementById('docSearch');
        const selectedDoc = document.querySelector('input[name="selectedDoc"]');
        if (!fileInput || !displayInput || !docSearch || !selectedDoc) return;
        if (fileInput._enhanced) return;
        fileInput._enhanced = true;

        // Default to Unknown Document until a file tells us otherwise
        applyDocType('', docSearch, selectedDoc);

        // Auto-fill display name from selected file (extension stripped) + set doc type
        fileInput.addEventListener('change', () => {
            const file = fileInput.files[0];
            if (!file) return;
            const baseName = file.name.replace(/\.[^/.]+$/, '');
            displayInput.value = baseName;
            applyDocType(baseName, docSearch, selectedDoc);
        });

        // If the user manually edits the display name, re-evaluate the doc type
        displayInput.addEventListener('input', () => {
            applyDocType(displayInput.value, docSearch, selectedDoc);
        });

        if (pendingFilePick) {
            pendingFilePick = false;
            fileInput.click();
        }
        hookUploadBtn();
    }
    const observer = new MutationObserver(() => {
        if (document.getElementById('selectedFile')) {
            enhanceModal();
        }
    });
    observer.observe(document.body, { childList: true, subtree: true });
    document.addEventListener('keydown', handleEnter, true);
    if (document.readyState === 'complete') {
        hookLoadDoc();
    } else {
        window.addEventListener('load', hookLoadDoc);
    }
})();