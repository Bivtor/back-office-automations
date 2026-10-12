// ==UserScript==
// @name         OES Assignment Assistant Upgrade
// @namespace    dq-tools
// @version      1.2
// @description  Paste the Daily Queue "modified" table (with header row); preview OES -> processor assignments (color-coded, deduped by OES) and apply them safely to the OES Assignment modal. Idempotent: never overwrites an existing assignment unless you explicitly choose to. Blocks apply if one OES maps to multiple processors. Mounts only while the "Assignment Administration" modal is open (SPA-aware via MutationObserver); hides (keeps your paste) when it closes.
// @match        https://oes.example.com/
// @match        https://oes.example.com/index.cfm?openModule=37
// @grant        none
// ==/UserScript==

(function () {
    'use strict';

    // -------------------------------------------------------------------------
    // CONFIG
    // -------------------------------------------------------------------------
    // Map the pipeline's Assignment-column names to the dropdown's numeric
    // option values. We match on VALUE (stable) rather than the dropdown's
    // display text, so name formatting differences ("D.Davis" vs "Davis,
    // Drew") never cause a mis-match.
    const NAME_TO_VALUE = {
        'f.foster':  '1001',  // Foster, Fran
        'v.rinaldi': '1002',  // Rinaldi, Victor
        'd.davis':   '1003',  // Davis, Drew
        'e.evans':   '1004',  // Evans, Evan
        'g.grant':   '1005'   // Grant, Gray  (PERSH)
    };
    const VALUE_TO_LABEL = {
        '1001': 'Foster, Fran',
        '1002': 'Rinaldi, Victor',
        '1003': 'Davis, Drew',
        '1004': 'Evans, Evan',
        '1005': 'Grant, Gray'
    };

    // The OES Assignment grid: select[name="assignment"].oesInput rows. We scope
    // strictly to these so other modals reusing the generic .modal wrappers are
    // never touched.
    const ROW_SELECTOR = 'select[name="assignment"].oesInput';
    // The modal is considered OPEN when a visible .modal-title reads this:
    const MODAL_TITLE_TEXT = 'assignment administration';

    // -------------------------------------------------------------------------
    // STATE  (persists across modal open/close; we only hide the panel)
    // -------------------------------------------------------------------------
    let parsedRows = [];   // [{ oes, procName, value, status, label, current, currentLabel, skip }]
    let lastPaste  = '';   // remember the textarea contents too
    let panelEl    = null; // the panel DOM node (built once, shown/hidden)

    // -------------------------------------------------------------------------
    // MODAL OPEN DETECTION
    // -------------------------------------------------------------------------
    function isElementVisible(el) {
        if (!el) return false;
        if (el.offsetParent === null && getComputedStyle(el).position !== 'fixed') {
            // offsetParent null usually means display:none somewhere up the tree.
            // (fixed elements have null offsetParent but can still be visible.)
            const rect = el.getBoundingClientRect();
            if (rect.width === 0 && rect.height === 0) return false;
        }
        const style = getComputedStyle(el);
        return style.display !== 'none' && style.visibility !== 'hidden';
    }

    // Returns the open Assignment Administration modal element, or null.
    function findOpenAssignmentModal() {
        const titles = document.querySelectorAll('.modal-title');
        for (const t of titles) {
            if (t.textContent.trim().toLowerCase() === MODAL_TITLE_TEXT) {
                // Walk up to the modal container and confirm it's visible.
                const modal = t.closest('.modal') || t.closest('.modal-content') || t.parentElement;
                if (modal && isElementVisible(modal)) {
                    // Extra guard: the assignment grid must actually be present.
                    if (modal.querySelector(ROW_SELECTOR)) return modal;
                    // Grid may live in a sibling tab container; accept the modal
                    // anyway if any oesInput exists in the document.
                    if (document.querySelector(ROW_SELECTOR)) return modal;
                }
            }
        }
        return null;
    }

    // -------------------------------------------------------------------------
    // PAGE READ HELPERS  (read-only)
    // -------------------------------------------------------------------------
    function indexPageSelects() {
        const map = {};
        document.querySelectorAll(ROW_SELECTOR).forEach(sel => {
            let oes = null;
            const oc = sel.getAttribute('onchange') || '';
            const m = oc.match(/updateAssignmentAdmin\([^,]+,\s*(\d+)\)/);
            if (m) oes = m[1];
            if (!oes) {
                const tr = sel.closest('tr');
                if (tr) { const td = tr.querySelector('td'); if (td) oes = td.textContent.trim(); }
            }
            if (!oes) return;
            let cur = '';
            const selOpt = sel.querySelector('option[selected]');
            if (selOpt) cur = (selOpt.getAttribute('value') || '').trim();
            else cur = (sel.value || '').trim();
            if (cur === '0') cur = '';
            map[oes] = { select: sel, currentValue: cur };
        });
        return map;
    }

    // -------------------------------------------------------------------------
    // PARSE PASTED TABLE
    // -------------------------------------------------------------------------
    // Parse the pasted "modified" sheet (tab-separated, WITH header row) and
    // collapse to one entry per unique OES.
    //
    // Multiple ROs can share an OES. Rules for resolving an OES's processor:
    //   - Ignore blank-assignment legs (internal/in-house rows have no proc).
    //   - Take the non-blank processor shared by the OES's other legs.
    //   - If two legs of the SAME OES name DIFFERENT non-blank processors, that
    //     is a hard error: the OES is flagged status 'error' and the caller
    //     disables Apply entirely until it's resolved upstream.
    //
    // Returns [{ oes, procName, error?:true, conflictProcs?:[...] }] preserving
    // first-seen OES order.
    function parsePaste(text) {
        if (!text || !text.trim()) return [];

        const lines = text.replace(/\r/g, '').split('\n').filter(l => l.trim() !== '');
        if (lines.length === 0) return [];
        const rows = lines.map(l => l.includes('\t') ? l.split('\t') : l.split(/\s{2,}/));

        // Locate header columns. Accept "OES ID"/"OES"/"OESID" and "Assignment".
        // Also find the "R/O ID" column so the fallback can EXCLUDE it (both RO
        // and OES are numeric, and RO comes first, so a naive "first number"
        // fallback would wrongly grab the RO).
        let oesCol = -1, asgCol = -1, roCol = -1, headerIdx = -1;
        for (let i = 0; i < rows.length; i++) {
            const cells = rows[i].map(c => c.trim().toLowerCase());
            const oi = cells.findIndex(c => c === 'oes id' || c === 'oes' || c === 'oesid');
            const ai = cells.findIndex(c => c === 'assignment');
            if (oi !== -1 && ai !== -1) {
                oesCol = oi; asgCol = ai; headerIdx = i;
                roCol = cells.findIndex(c => c === 'r/o id' || c === 'ro id' || c === 'r/o' || c === 'roid');
                break;
            }
        }

        const knownNames = Object.keys(NAME_TO_VALUE);

        // First pass: pull (oes, proc) from every data row.
        const raw = [];
        for (let i = 0; i < rows.length; i++) {
            if (i === headerIdx) continue;
            const cells = rows[i].map(c => c.trim());

            let oes = '', proc = '';
            if (oesCol !== -1 && asgCol !== -1) {
                oes  = (cells[oesCol] || '').trim();
                proc = (cells[asgCol] || '').trim();
            } else {
                // Fallback (no header found): take the SECOND numeric cell as OES
                // when a first numeric (RO) precedes it; otherwise the first.
                const nums = [];
                cells.forEach((c, idx) => { if (/^\d{4,}$/.test(c)) nums.push({ idx, val: c }); });
                if (nums.length >= 2) oes = nums[1].val;       // RO, OES, ... -> OES is 2nd
                else if (nums.length === 1) oes = nums[0].val;
                proc = (cells.find(c => knownNames.includes(c.toLowerCase())) || '').trim();
            }

            if (!/^\d{4,}$/.test(oes)) continue;  // skip separators/blank/no-OES rows
            raw.push({ oes: oes, proc: proc });   // proc may be '' (blank leg) -- kept for now
        }

        // Second pass: group by OES, resolve processor, detect conflicts.
        const order = [];
        const byOes = {};
        raw.forEach(r => {
            if (!byOes[r.oes]) { byOes[r.oes] = { procs: new Set(), any: false }; order.push(r.oes); }
            byOes[r.oes].any = true;
            if (r.proc) byOes[r.oes].procs.add(r.proc);   // only collect non-blank procs
        });

        const out = [];
        order.forEach(oes => {
            const procs = Array.from(byOes[oes].procs);
            if (procs.length === 0) return;                 // all legs blank -> nothing to assign
            if (procs.length === 1) {
                out.push({ oes: oes, procName: procs[0] });
            } else {
                // Same OES, multiple DIFFERENT non-blank processors -> hard error.
                out.push({ oes: oes, procName: procs[0], error: true, conflictProcs: procs });
            }
        });
        return out;
    }

    // -------------------------------------------------------------------------
    // CLASSIFY  (read-only against the live page)
    // -------------------------------------------------------------------------
    function classify(rows) {
        const page = indexPageSelects();
        return rows.map(r => {
            const value = NAME_TO_VALUE[r.procName.toLowerCase()] || null;
            const rec = {
                oes: r.oes, procName: r.procName, value: value,
                label: value ? (VALUE_TO_LABEL[value] || r.procName) : r.procName,
                current: '', currentLabel: '', status: 'gray', skip: false,
                conflictProcs: r.conflictProcs || null
            };

            // A parse-level conflict (same OES -> different processors) is a hard
            // error and outranks every page-based status. It disables Apply.
            if (r.error) { rec.status = 'error'; return rec; }

            if (!value) { rec.status = 'bad'; return rec; }

            const onPage = page[r.oes];
            if (!onPage) { rec.status = 'gray'; return rec; }

            rec.current = onPage.currentValue;
            rec.currentLabel = onPage.currentValue ? (VALUE_TO_LABEL[onPage.currentValue] || onPage.currentValue) : '';

            if (!onPage.currentValue)                 rec.status = 'green';
            else if (onPage.currentValue === value)   rec.status = 'yellow';
            else                                      rec.status = 'red';
            return rec;
        });
    }

    // Re-run classification against current DOM, preserving each row's skip flag.
    function reclassifyPreservingSkips() {
        if (parsedRows.length === 0) return;
        const skips = parsedRows.map(r => r.skip);
        parsedRows = classify(parsedRows.map(r => ({
            oes: r.oes, procName: r.procName,
            error: r.status === 'error' ? true : undefined,
            conflictProcs: r.conflictProcs || undefined
        })));
        parsedRows.forEach((r, i) => { r.skip = skips[i] || false; });
    }

    // -------------------------------------------------------------------------
    // APPLY  (only mutation point)
    // -------------------------------------------------------------------------
    function applyAssignments(allowRed) {
        const page = indexPageSelects();
        let applied = 0, skipped = 0, failed = 0;
        const details = [];

        parsedRows.forEach(rec => {
            if (rec.skip) { skipped++; return; }
            if (rec.status === 'error' || rec.status === 'bad' || rec.status === 'gray') { skipped++; return; }
            if (rec.status === 'yellow') { skipped++; return; }
            if (rec.status === 'red' && !allowRed) { skipped++; return; }

            const onPage = page[rec.oes];
            if (!onPage || !rec.value) { failed++; return; }

            try {
                onPage.select.value = rec.value;
                if (typeof window.updateAssignmentAdmin === 'function') {
                    window.updateAssignmentAdmin(rec.value, rec.oes);
                } else {
                    onPage.select.dispatchEvent(new Event('change', { bubbles: true }));
                }
                applied++;
                details.push(`${rec.oes} -> ${rec.label}${rec.status === 'red' ? ' (OVERWROTE ' + rec.currentLabel + ')' : ''}`);
            } catch (e) {
                failed++;
            }
        });

        return { applied, skipped, failed, details };
    }

    // -------------------------------------------------------------------------
    // UI
    // -------------------------------------------------------------------------
    const COLORS = {
        green:  { bg: '#e6f4ea', border: '#34a853', label: 'OK (blank -> assign)' },
        yellow: { bg: '#fef7e0', border: '#f9ab00', label: 'Already assigned (same)' },
        red:    { bg: '#fce8e6', border: '#ea4335', label: 'CONFLICT (assigned, different)' },
        error:  { bg: '#5b0000', border: '#5b0000', label: 'ERROR (OES has 2+ processors)' },
        gray:   { bg: '#f1f3f4', border: '#9aa0a6', label: 'Not on this page' },
        bad:    { bg: '#f1f3f4', border: '#9aa0a6', label: 'Unknown processor' }
    };

    function buildPanel() {
        if (panelEl) return panelEl;

        const panel = document.createElement('div');
        panel.id = 'dqApplierPanel';
        panel.style.cssText = [
            'position:fixed', 'top:60px', 'right:16px', 'width:440px', 'max-height:80vh',
            'background:#fff', 'border:1px solid #ccc', 'border-radius:8px',
            'box-shadow:0 4px 16px rgba(0,0,0,.2)', 'z-index:2147483647',
            'font-family:Segoe UI,Arial,sans-serif', 'font-size:13px',
            'display:none', 'flex-direction:column', 'overflow:hidden'
        ].join(';');

        panel.innerHTML = `
            <div style="background:#1f3a5f;color:#fff;padding:8px 12px;font-weight:600;display:flex;justify-content:space-between;align-items:center;cursor:move" id="dqHeader">
                <span>DQ Assignment Applier</span>
                <span style="cursor:pointer;font-size:16px;line-height:1" id="dqClose" title="Hide">&times;</span>
            </div>
            <div style="padding:10px 12px;overflow:auto">
                <div style="margin-bottom:6px;color:#444">Paste the <b>modified</b> sheet (include the header row):</div>
                <textarea id="dqPaste" rows="5" style="width:100%;box-sizing:border-box;font-family:monospace;font-size:11px;resize:vertical"></textarea>
                <div style="display:flex;gap:8px;margin:8px 0">
                    <button id="dqParse" style="flex:1;padding:6px;cursor:pointer">Preview</button>
                    <button id="dqClear" style="padding:6px 10px;cursor:pointer">Clear</button>
                </div>
                <div id="dqLegend" style="font-size:11px;color:#555;margin-bottom:6px"></div>
                <div id="dqSummary" style="font-size:12px;font-weight:600;margin-bottom:6px"></div>
                <div id="dqList" style="max-height:34vh;overflow:auto;border:1px solid #eee;border-radius:4px"></div>
                <label style="display:flex;align-items:center;gap:6px;margin:8px 0;font-size:12px;color:#a00">
                    <input type="checkbox" id="dqAllowRed"> Also overwrite RED conflicts (changes an existing assignment)
                </label>
                <button id="dqApply" style="width:100%;padding:8px;background:#34a853;color:#fff;border:none;border-radius:4px;font-weight:600;cursor:pointer" disabled>Apply enabled assignments</button>
                <div id="dqResult" style="font-size:12px;margin-top:8px;white-space:pre-wrap"></div>
            </div>
        `;
        document.body.appendChild(panel);
        panelEl = panel;

        panel.querySelector('#dqLegend').innerHTML = Object.keys(COLORS).map(k =>
                                                                             `<span style="display:inline-block;margin-right:8px"><span style="display:inline-block;width:10px;height:10px;background:${COLORS[k].bg};border:1px solid ${COLORS[k].border};vertical-align:middle;margin-right:3px"></span>${COLORS[k].label}</span>`
                                                                            ).join('');

        panel.querySelector('#dqClose').onclick = () => { panel.style.display = 'none'; };
        panel.querySelector('#dqClear').onclick = () => {
            panel.querySelector('#dqPaste').value = '';
            lastPaste = ''; parsedRows = []; renderList();
        };
        panel.querySelector('#dqPaste').addEventListener('input', e => { lastPaste = e.target.value; });
        panel.querySelector('#dqParse').onclick = () => {
            lastPaste = panel.querySelector('#dqPaste').value;
            parsedRows = classify(parsePaste(lastPaste));
            renderList();
        };
        panel.querySelector('#dqApply').onclick = () => {
            const errorTotal = parsedRows.filter(r => r.status === 'error').length;
            if (errorTotal > 0) {
                alert(`Cannot apply: ${errorTotal} OES assigned to multiple processors. Fix the source data and re-paste.`);
                return;
            }
            const allowRed = panel.querySelector('#dqAllowRed').checked;
            const greenCount = parsedRows.filter(r => !r.skip && r.status === 'green').length;
            const redCount   = parsedRows.filter(r => !r.skip && r.status === 'red').length;
            let msg = `Apply ${greenCount} new assignment(s)?`;
            if (allowRed && redCount > 0) msg += `\n\nThis will ALSO OVERWRITE ${redCount} existing (RED) assignment(s).`;
            if (greenCount === 0 && !(allowRed && redCount > 0)) { alert('Nothing to apply.'); return; }
            if (!confirm(msg)) return;
            const res = applyAssignments(allowRed);
            const result = panel.querySelector('#dqResult');
            result.style.color = res.failed > 0 ? '#a00' : '#137333';
            result.textContent = `Applied: ${res.applied}   Skipped: ${res.skipped}   Failed: ${res.failed}\n` +
                (res.details.length ? res.details.join('\n') : '');
            reclassifyPreservingSkips();
            renderList();
            lastGridSig = gridSignature();
        };

        makeDraggable(panel, panel.querySelector('#dqHeader'));
        return panel;
    }

    function renderList() {
        if (!panelEl) return;
        const list = panelEl.querySelector('#dqList');
        const summary = panelEl.querySelector('#dqSummary');
        const applyBtn = panelEl.querySelector('#dqApply');
        const paste = panelEl.querySelector('#dqPaste');
        if (paste && paste.value !== lastPaste) paste.value = lastPaste;

        if (parsedRows.length === 0) {
            list.innerHTML = '<div style="padding:10px;color:#888">No rows parsed yet.</div>';
            summary.textContent = '';
            applyBtn.disabled = true;
            return;
        }

        const counts = { green: 0, yellow: 0, red: 0, gray: 0, bad: 0, error: 0 };
        parsedRows.forEach(r => { if (!r.skip) counts[r.status]++; });
        // Errors count regardless of skip: a parse conflict means the pasted data
        // is bad and must be fixed upstream, so we never let Apply run.
        const errorTotal = parsedRows.filter(r => r.status === 'error').length;

        summary.innerHTML =
            `<span style="color:#137333">${counts.green} new</span> &middot; ` +
            `<span style="color:#b06000">${counts.yellow} same</span> &middot; ` +
            `<span style="color:#c5221f">${counts.red} conflict</span> &middot; ` +
            `<span style="color:#666">${counts.gray + counts.bad} skip</span>` +
            (errorTotal > 0 ? ` &middot; <span style="color:#5b0000;font-weight:700">${errorTotal} ERROR</span>` : '');

        // Hard-error banner: blocks Apply until the upstream paste is fixed.
        if (errorTotal > 0) {
            const errs = parsedRows.filter(r => r.status === 'error')
            .map(r => `OES ${r.oes}: ${(r.conflictProcs || []).join(' vs ')}`);
            summary.innerHTML +=
                `<div style="margin-top:6px;padding:6px 8px;background:#5b0000;color:#fff;border-radius:4px;font-weight:600">` +
                `Cannot apply &mdash; ${errorTotal} OES assigned to multiple processors:<br>` +
                errs.map(e => `&bull; ${e}`).join('<br>') +
                `<div style="font-weight:400;margin-top:4px">Fix the source data and re-paste.</div></div>`;
        }

        list.innerHTML = '';
        parsedRows.forEach((r, idx) => {
            const c = COLORS[r.status];
            const row = document.createElement('div');
            row.style.cssText = `display:flex;align-items:center;gap:8px;padding:5px 8px;border-bottom:1px solid #f0f0f0;background:${r.skip ? '#fafafa' : c.bg};${r.status === 'error' ? 'color:#fff;' : ''}${r.skip ? 'opacity:.5;text-decoration:line-through' : ''}`;
            const detail = r.status === 'error'
            ? `<b>2+ processors:</b> ${(r.conflictProcs || []).join(', ')}`
            : (r.status === 'red'
               ? `<span style="color:#c5221f">now: ${r.currentLabel}</span> &rarr; <b>${r.label}</b>`
               : (r.status === 'gray' ? `<i>not on page</i>`
                  : (r.status === 'bad' ? `<i>unknown: ${r.procName}</i>`
                     : `&rarr; <b>${r.label}</b>`)));
            const xColor = r.status === 'error' ? '#fff' : '#999';
            row.innerHTML = `
                <span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:${c.border};border:1px solid ${r.status === 'error' ? '#fff' : c.border};flex:0 0 auto"></span>
                <span style="flex:0 0 auto;font-family:monospace">${r.oes}</span>
                <span style="flex:1">${detail}</span>
                <span data-idx="${idx}" class="dqX" style="cursor:pointer;color:${xColor};font-weight:700;flex:0 0 auto" title="Toggle include/exclude">&times;</span>
            `;
            list.appendChild(row);
        });

        list.querySelectorAll('.dqX').forEach(x => {
            x.onclick = () => {
                const i = parseInt(x.getAttribute('data-idx'), 10);
                if (!isNaN(i) && parsedRows[i]) { parsedRows[i].skip = !parsedRows[i].skip; renderList(); }
            };
        });

        // Disable Apply if there's nothing to do OR any hard error is present.
        applyBtn.disabled = errorTotal > 0 || (counts.green === 0 && counts.red === 0);
        applyBtn.title = errorTotal > 0 ? 'Resolve OES-with-multiple-processors errors first' : '';
    }

    function makeDraggable(el, handle) {
        let ox = 0, oy = 0, dragging = false;
        handle.addEventListener('mousedown', e => {
            dragging = true; ox = e.clientX - el.offsetLeft; oy = e.clientY - el.offsetTop; e.preventDefault();
        });
        document.addEventListener('mousemove', e => {
            if (!dragging) return;
            el.style.left = (e.clientX - ox) + 'px';
            el.style.top  = (e.clientY - oy) + 'px';
            el.style.right = 'auto';
        });
        document.addEventListener('mouseup', () => { dragging = false; });
    }

    // -------------------------------------------------------------------------
    // LIFECYCLE  (mount-on-open / hide-on-close, SPA-aware)
    // -------------------------------------------------------------------------
    let modalWasOpen = false;
    let lastGridSig  = '';   // cheap signature of the grid to detect real changes

    // Cheap fingerprint of the current grid: count of rows + concatenated OES ids.
    // Lets us skip re-classification when nothing about the grid changed.
    function gridSignature() {
        const sels = document.querySelectorAll(ROW_SELECTOR);
        let sig = sels.length + '|';
        sels.forEach(s => {
            const oc = s.getAttribute('onchange') || '';
            const m = oc.match(/,\s*(\d+)\)/);
            sig += (m ? m[1] : '') + ',' + (s.value || '') + ';';
        });
        return sig;
    }

    function syncToModal() {
        const open = !!findOpenAssignmentModal();

        if (open && !modalWasOpen) {
            // Modal just opened: ensure panel exists, restore paste, re-classify
            // against the freshly-rendered grid, then show.
            buildPanel();
            const paste = panelEl.querySelector('#dqPaste');
            if (paste) paste.value = lastPaste;
            reclassifyPreservingSkips();
            renderList();
            panelEl.style.display = 'flex';
            lastGridSig = gridSignature();
        } else if (open && modalWasOpen) {
            // Still open: only re-classify if the grid actually changed
            // (e.g. an assignment we applied updated a selected option).
            const sig = gridSignature();
            if (sig !== lastGridSig) {
                lastGridSig = sig;
                reclassifyPreservingSkips();
                renderList();
            }
        } else if (!open && modalWasOpen) {
            // Modal just closed: hide but keep state.
            if (panelEl) panelEl.style.display = 'none';
            lastGridSig = '';
        }

        modalWasOpen = open;
    }

    // Coalesce ExtJS/Bootstrap mutation bursts into one sync per frame.
    let scheduled = false;
    const observer = new MutationObserver(() => {
        if (scheduled) return;
        scheduled = true;
        requestAnimationFrame(() => { scheduled = false; syncToModal(); });
    });
    observer.observe(document.body, { childList: true, subtree: true, attributes: true, attributeFilter: ['class', 'style'] });

    syncToModal();
})();