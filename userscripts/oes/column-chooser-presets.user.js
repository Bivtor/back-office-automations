// ==UserScript==
// @name         OES Column Chooser Preset
// @namespace    oes-tools
// @version      1.1.0
// @description  Adds preset buttons that apply fixed column sets via the ColVis Column Chooser
// @match        https://oes.example.com/*
// @match        https://partners.example.com/*
// @match        https://*.partners.example.com/*
// @grant        none
// @run-at       document-idle
// ==/UserScript==

(function () {
    'use strict';

    // Each preset: exact columns to show. Everything else in the chooser gets unchecked.
    var PRESETS = [
        {
            id: 'tm-colpreset-weekly',
            label: 'Weekly Report',
            columns: [
                'Client Name',
                'Source Company',
                'Created',
                'LOA Date',
                'TPA Authorization Requested',
                'TPA Approval',
                'Expected Amount',
                'Source Account',
                'OES Assigned To',
                'OES',
                'Representative',
                'Rollover ID',
                'Total Platform Balance'
            ]
        },
        {
            id: 'tm-colpreset-outstanding-ro',
            label: 'Outstanding RO Report',
            columns: [
                'Source Company',
                'Client Name',
                'Source Account',
                'OES',
                'Rollover ID',
                'Representative',
                'LOA Date',
                'TPA Authorization Requested',
                'TPA Approval',
                'Expected Amount',
                'Received Amount',
                'Target Account',
                'Target Product Type',
                'Target Product',
                'Status',
                'Client Initiated',
                'Internal',
                'Funded at a Later Date'
            ]
        }
    ];

    var busy = false;

    // ---------- helpers ----------

    function btnHtml(label) {
        return '<i class="fa fa-columns" style="margin-right:10px;"></i>' + label;
    }

    function isVisible(el) {
        return !!el && el.isConnected && el.getClientRects().length > 0 &&
            getComputedStyle(el).display !== 'none';
    }

    function sleep(ms) {
        return new Promise(function (r) { setTimeout(r, ms); });
    }

    async function waitFor(fn, timeoutMs) {
        var start = Date.now();
        while (Date.now() - start < timeoutMs) {
            var v = fn();
            if (v) return v;
            await sleep(50);
        }
        return null;
    }

    function getCollection(requireVisible) {
        var lists = Array.prototype.slice.call(
            document.querySelectorAll('ul.ColVis_collection')
        ).filter(function (ul) { return ul.querySelector('input.ColVisSel'); });

        var vis = lists.filter(isVisible);
        if (vis.length) return vis[vis.length - 1];
        return requireVisible ? null : (lists[lists.length - 1] || null);
    }

    function boxName(box) {
        var v = (box.value || '').trim();
        if (v) return v;
        var span = box.parentElement && box.parentElement.querySelector('span');
        return span ? span.textContent.trim() : '';
    }

    function findBox(ul, name) {
        var boxes = ul.querySelectorAll('input.ColVisSel');
        for (var i = 0; i < boxes.length; i++) {
            if (boxName(boxes[i]) === name) return boxes[i];
        }
        return null;
    }

    function flash(btn, preset, text, isError) {
        btn.innerHTML = text;
        btn.title = text;
        if (isError) btn.classList.add('text-danger');
        setTimeout(function () {
            btn.innerHTML = btnHtml(preset.label);
            btn.title = 'Apply ' + preset.label + ' columns';
            btn.classList.remove('text-danger');
        }, isError ? 4000 : 1500);
    }

    function setAllDisabled(disabled) {
        PRESETS.forEach(function (p) {
            var b = document.getElementById(p.id);
            if (b) b.disabled = disabled;
        });
    }

    // ---------- core ----------

    async function applyPreset(btn, preset, colVisDiv) {
        if (busy) return;
        busy = true;
        setAllDisabled(true);

        var openedByUs = false;

        try {
            var ul = getCollection(true);

            // Chooser list not open: open it via the Column Chooser button
            if (!ul) {
                var chooserBtn = colVisDiv.querySelector('button');
                if (!chooserBtn) throw new Error('Column Chooser button not found');
                chooserBtn.click();
                openedByUs = true;
                ul = await waitFor(function () { return getCollection(true); }, 3000);
            }

            if (!ul) throw new Error('Column list did not open');

            var want = {};
            preset.columns.forEach(function (n) { want[n] = true; });

            var names = Array.prototype.map.call(
                ul.querySelectorAll('input.ColVisSel'), boxName
            );

            var missing = preset.columns.filter(function (n) { return names.indexOf(n) === -1; });

            // Re-query each box every pass in case the list gets rebuilt after a toggle
            for (var i = 0; i < names.length; i++) {
                var name = names[i];
                var current = getCollection(false) || ul;
                var box = findBox(current, name);
                if (!box) continue;

                var shouldBeChecked = !!want[name];
                if (box.checked !== shouldBeChecked) {
                    box.click();
                    await sleep(0);
                }
            }

            if (openedByUs) {
                await sleep(100);
                var bg = document.querySelector('.ColVis_collectionBackground');
                if (bg) bg.click();
            }

            if (missing.length) {
                flash(btn, preset, 'Missing: ' + missing.join(', '), true);
                console.warn('[ColPreset] ' + preset.label + ' - columns not found in chooser:', missing);
            } else {
                flash(btn, preset, '<i class="fa fa-check" style="margin-right:10px;"></i>Applied');
            }
        } catch (err) {
            console.error('[ColPreset]', err);
            flash(btn, preset, 'Error: ' + err.message, true);
        } finally {
            busy = false;
            setAllDisabled(false);
        }
    }

    // ---------- injection ----------

    function inject() {
        var searchBtn = document.getElementById('SearchBtn');
        if (!searchBtn) return;

        var row = searchBtn.closest('.row');
        if (!row) return;

        var colVisDiv = row.querySelector('.ColVis');
        if (!colVisDiv) return;

        // Insert inline after Reset (falls back to after Search), in preset order
        var anchor = document.getElementById('ResetBtn') || searchBtn;

        PRESETS.forEach(function (preset) {
            var existing = row.querySelector('#' + preset.id);
            if (existing) {
                anchor = existing;
                return;
            }

            var btn = document.createElement('button');
            btn.type = 'button';
            btn.id = preset.id;
            btn.className = 'btn btn-light border';
            btn.style.marginRight = '10px';
            btn.title = 'Apply ' + preset.label + ' columns';
            btn.innerHTML = btnHtml(preset.label);

            btn.addEventListener('click', function (e) {
                e.preventDefault();
                e.stopPropagation();
                applyPreset(btn, preset, colVisDiv);
            });

            anchor.insertAdjacentElement('afterend', btn);
            anchor = btn;
        });
    }

    // SPA: watch for the row to appear / be re-rendered
    var scheduled = false;
    var observer = new MutationObserver(function () {
        if (scheduled) return;
        scheduled = true;
        setTimeout(function () {
            scheduled = false;
            inject();
        }, 250);
    });

    observer.observe(document.body, { childList: true, subtree: true });
    inject();
})();