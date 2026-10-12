// ==UserScript==
// @name         Auto JB Stamp
// @namespace    http://tampermonkey.net/
// @version      0.1.0
// @description  Auto-stamps EQ/JB PDF forms based on filename. Draws red X on page 8 and saves as _stamped.
// @author       You
// @match        file:///*
// @grant        GM_download
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_deleteValue
// @grant        GM_xmlhttpRequest
// @connect      localhost
// @require      https://unpkg.com/pdf-lib@1.17.1/dist/pdf-lib.min.js
// @run-at       document-start
// ==/UserScript==

(async function () {
    'use strict';

    // ─── DEBUG LOGGER ────────────────────────────────────────────────────────────
    const LOG_PREFIX = '[AutoJBStamp]';
    function log(...args)  { console.log  (LOG_PREFIX, ...args); }
    function warn(...args) { console.warn (LOG_PREFIX, ...args); }
    function err(...args)  { console.error(LOG_PREFIX, ...args); }

    log('Script injected. Current URL:', window.location.href);
    log('User agent:', navigator.userAgent);
    log('Document readyState:', document.readyState);

    // ─── URL PARSING ─────────────────────────────────────────────────────────────
    const rawUrl  = window.location.href;
    const fileUrl = decodeURIComponent(rawUrl);

    log('Decoded URL:', fileUrl);

    // Extract just the filename (no path, no extension)
    const filenameMatch = fileUrl.match(/([^/\\]+)\.pdf$/i);
    if (!filenameMatch) {
        log('URL does not end in .pdf — script will not run.');
        return;
    }
    const filename = filenameMatch[1]; // e.g. "eq_123456_A"
    log('Detected PDF filename (no ext):', filename);

    // ─── FILENAME PATTERN CHECK ──────────────────────────────────────────────────
    // Expected format: eq_[account#]_[boxSelection]
    const patternMatch = filename.match(/^eq_([^_]+)_([^_]+)$/i);
    if (!patternMatch) {
        log('Filename does not match eq_[account]_[box] pattern — script will not run.');
        log('Pattern tested against:', filename);
        return;
    }

    const accountNumber = patternMatch[1];
    const boxSelection  = patternMatch[2];
    log('✅ Pattern matched!');
    log('   Account # :', accountNumber);
    log('   Box select:', boxSelection);

    // ─── INFINITE LOOP GUARD (localStorage) ──────────────────────────────────────
    // Key includes the filename so different files still get processed.
    // Value is a timestamp; we consider it "done" if it was set within the last 60s.
    const guardKey = `AutoJBStamp_processed_${filename}`;
    const GUARD_TTL_MS = 60_000; // 60 seconds

    const existingGuard = localStorage.getItem(guardKey);
    if (existingGuard) {
        const age = Date.now() - parseInt(existingGuard, 10);
        if (age < GUARD_TTL_MS) {
            log(`Loop guard active — this file was already processed ${age}ms ago. Aborting to prevent infinite loop.`);
            return;
        } else {
            log(`Loop guard expired (was set ${age}ms ago). Clearing and re-processing.`);
            localStorage.removeItem(guardKey);
        }
    }

    log('Loop guard: not set or expired. Proceeding with stamp operation.');

    // ─── WAIT FOR DOCUMENT LOAD ───────────────────────────────────────────────────
    // For file:/// PDFs the browser renders the PDF natively; the "document" is
    // essentially an embed. We just need the page to be ready enough to fetch.
    await new Promise(resolve => {
        if (document.readyState === 'complete') {
            log('Document already complete.');
            resolve();
        } else {
            log('Waiting for document load...');
            window.addEventListener('load', () => {
                log('Document load event fired.');
                resolve();
            });
        }
    });

    // ─── FETCH THE PDF BYTES (via GM_xmlhttpRequest to bypass CORS on file:///) ──
    // fetch() is blocked by Edge CORS policy for file:/// origins.
    // GM_xmlhttpRequest runs in the extension context and is not subject to this.
    log('Fetching PDF via GM_xmlhttpRequest from:', rawUrl);
    let pdfBytes;
    try {
        pdfBytes = await new Promise((resolve, reject) => {
            GM_xmlhttpRequest({
                method:       'GET',
                url:          rawUrl,
                responseType: 'arraybuffer',
                onload: (response) => {
                    log('GM_xmlhttpRequest onload. Status:', response.status, '| Byte length:', response.response.byteLength);
                    if (response.status === 0 || response.status === 200) {
                        // status 0 is normal for file:// requests — not an error
                        resolve(response.response);
                    } else {
                        reject(new Error(`Unexpected status: ${response.status}`));
                    }
                },
                onerror: (response) => {
                    reject(new Error(`GM_xmlhttpRequest error: ${JSON.stringify(response)}`));
                },
                ontimeout: () => {
                    reject(new Error('GM_xmlhttpRequest timed out'));
                },
            });
        });
        log('PDF fetched successfully. Byte length:', pdfBytes.byteLength);
    } catch (e) {
        err('Failed to fetch PDF:', e);
        return;
    }

    // ─── LOAD WITH pdf-lib ────────────────────────────────────────────────────────
    log('Checking pdf-lib availability:', typeof PDFLib);
    if (typeof PDFLib === 'undefined') {
        err('pdf-lib is NOT loaded! Check that @require CDN URL is reachable.');
        return;
    }

    const { PDFDocument, rgb, degrees } = PDFLib;

    let pdfDoc;
    try {
        pdfDoc = await PDFDocument.load(pdfBytes);
        log('PDF loaded by pdf-lib. Total pages:', pdfDoc.getPageCount());
    } catch (e) {
        err('pdf-lib failed to parse PDF:', e);
        return;
    }

    // ─── TARGET PAGE 8 (index 7) ──────────────────────────────────────────────────
    const PAGE_INDEX = 7; // 0-based → page 8
    const totalPages = pdfDoc.getPageCount();

    if (totalPages < 8) {
        warn(`PDF only has ${totalPages} page(s); cannot target page 8. Will use last page instead.`);
    }

    const targetIndex = Math.min(PAGE_INDEX, totalPages - 1);
    const page = pdfDoc.getPage(targetIndex);
    const { width, height } = page.getSize();
    log(`Targeting page index ${targetIndex} (page ${targetIndex + 1}). Size: ${width} x ${height}`);

    // ─── DRAW RED X IN CENTER OF PAGE ────────────────────────────────────────────
    // Two diagonal lines crossing at center
    const cx = width  / 2;
    const cy = height / 2;
    const ARM = 40; // half-length of each X arm in points
    const LINE_WIDTH = 6;
    const RED = rgb(1, 0, 0);

    log(`Drawing red X at center (${cx}, ${cy}) with arm length ${ARM}.`);

    // pdf-lib uses a drawing operators approach via page.drawLine
    page.drawLine({
        start: { x: cx - ARM, y: cy - ARM },
        end:   { x: cx + ARM, y: cy + ARM },
        thickness: LINE_WIDTH,
        color: RED,
    });

    page.drawLine({
        start: { x: cx + ARM, y: cy - ARM },
        end:   { x: cx - ARM, y: cy + ARM },
        thickness: LINE_WIDTH,
        color: RED,
    });

    log('Red X drawn successfully.');

    // ─── SERIALIZE MODIFIED PDF ───────────────────────────────────────────────────
    log('Serializing modified PDF...');
    let modifiedBytes;
    try {
        modifiedBytes = await pdfDoc.save();
        log('Serialized. Output byte length:', modifiedBytes.byteLength);
    } catch (e) {
        err('Failed to serialize PDF:', e);
        return;
    }

    // ─── SET LOOP GUARD BEFORE TRIGGERING DOWNLOAD ───────────────────────────────
    // Set this BEFORE the download so that if the download causes any reload,
    // the guard is already in place.
    localStorage.setItem(guardKey, String(Date.now()));
    log('Loop guard set in localStorage with key:', guardKey);

    // ─── TRIGGER DOWNLOAD ────────────────────────────────────────────────────────
    const outputFilename = `${filename}_stamped.pdf`;
    log('Triggering download as:', outputFilename);

    try {
        const blob = new Blob([modifiedBytes], { type: 'application/pdf' });
        const blobUrl = URL.createObjectURL(blob);

        const anchor = document.createElement('a');
        anchor.href = blobUrl;
        anchor.download = outputFilename;
        anchor.style.display = 'none';
        document.body.appendChild(anchor);
        anchor.click();

        // Clean up
        setTimeout(() => {
            document.body.removeChild(anchor);
            URL.revokeObjectURL(blobUrl);
            log('Blob URL revoked and anchor removed.');
        }, 5000);

        log('✅ Download triggered successfully for:', outputFilename);
    } catch (e) {
        err('Failed to trigger download:', e);
        // Clear the guard if download failed so user can retry
        localStorage.removeItem(guardKey);
        return;
    }

    log('=== Auto JB Stamp complete ===');
    log('  Account #     :', accountNumber);
    log('  Box selection :', boxSelection);
    log('  Output file   :', outputFilename);
    log('  Page stamped  : page', targetIndex + 1);

})();