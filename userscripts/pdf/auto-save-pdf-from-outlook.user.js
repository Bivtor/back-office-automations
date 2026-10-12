// ==UserScript==
// @name         Auto-Save PDF to dl folder
// @namespace    http://tampermonkey.net/
// @version      1.0
// @description  Auto-downloads PDFs opened from Outlook cache with standardized datetime filename
// @author       Victor
// @match        file:///C:/Users/USERNAME/AppData/Local/Microsoft/Windows/INetCache/Content.Outlook/*/*
// @downloadURL  https://gist.githubusercontent.com/youruser/yourgist/raw/script.user.js
// @updateURL    https://gist.githubusercontent.com/youruser/yourgist/raw/script.user.js
// @grant        none
// ==/UserScript==

(function () {
    'use strict';


    console.log('Script fired');

    if (localStorage.getItem('autoDownloadDone')) return;
    localStorage.setItem('autoDownloadDone', 'true');

    // clear it after 3 seconds so next PDF works fine
    setTimeout(() => localStorage.removeItem('autoDownloadDone'), 3000);

    function getDateTimeFilename() {
        const now = new Date();
        const yyyy = now.getFullYear();
        const mm = String(now.getMonth() + 1).padStart(2, '0');
        const dd = String(now.getDate()).padStart(2, '0');
        const hh = String(now.getHours()).padStart(2, '0');
        const min = String(now.getMinutes()).padStart(2, '00');
        const ss = String(now.getSeconds()).padStart(2, '0');
        return `${yyyy}-${mm}-${dd}_${hh}${min}${ss}.pdf`;
    }

    function triggerDownload(url, filename) {
        console.log('Downloading as:', filename);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    }

    const currentURL = window.location.href.split('#')[0];

    setTimeout(() => {
        console.log('timeout fired');
        const currentURL = window.location.href.split('#')[0];
        if (currentURL.toLowerCase().endsWith('.pdf')) {
            triggerDownload(currentURL, getDateTimeFilename());
        }
    }, 1500);

})();