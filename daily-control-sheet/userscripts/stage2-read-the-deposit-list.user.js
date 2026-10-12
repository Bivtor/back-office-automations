// ==UserScript==
// @name         Control Sheet — remittance upload, stage 2 (read the deposit list)
// @namespace    control-sheet
// @version      1.0
// @description  On the Daily Deposit Details list, read every deposit, match it against the control sheet's figures, and show what the automation would do. Reviews only — clicks nothing.
// @match        https://www.partners.example.com/SimpleAlloc/DailyDeposit/DailyDepositDetails.cfm*
// @grant        GM_xmlhttpRequest
// @grant        unsafeWindow
// @connect      127.0.0.1
// @connect      localhost
// @run-at       document-idle
// ==/UserScript==

/*
 * This stage reads and proposes. It does not delete, split, or upload.
 *
 * That split is deliberate rather than unfinished. The list gives one handle
 * per row — the amount — and no sender or reference, so every match here is an
 * inference. Deleting a deposit is not reversible from the page in any obvious
 * way, and the amounts genuinely collide: today's list has $900.00 three times
 * and $150.00 twice. So the plan gets shown and you approve it; the acting
 * comes later, once the rules for what to delete are written down.
 */

(function () {
  'use strict';

  const SERVER = 'http://127.0.0.1:8777';

  const ask = (path, method = 'GET', payload = null) => new Promise((res, rej) => {
    GM_xmlhttpRequest({
      method, url: SERVER + path, timeout: 6000,
      headers: payload ? {'Content-Type': 'application/json'} : {},
      data: payload ? JSON.stringify(payload) : undefined,
      onload: r => { try { res(JSON.parse(r.responseText)); } catch (e) { rej(e); } },
      onerror: rej, ontimeout: rej,
    });
  });

  const PACE = 850;
  const sleep = ms => new Promise(r => setTimeout(r, ms));

  const cash = t => {
    const m = String(t || '').replace(/[^0-9.\-]/g, '');
    return m === '' ? null : parseFloat(m);
  };
  const usd = n => n == null ? '—'
    : '$' + n.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2});

  /* ── reading the table ──────────────────────────────────── */
  /* Columns, left to right: id, type, three empties, amount, empty, the
     details button, an editable amount, the total, the split cell, the
     buttons. Read by position because none of them carry a class or a header
     that survives the DataTable. */
  function readRows() {
    const out = [];
    for (const tr of document.querySelectorAll('tr[id^="row-"]')) {
      const cells = [...tr.cells].map(td => td.textContent.trim());
      out.push({
        id: tr.id.replace('row-', ''),
        type: cells[1] || '',
        amount: cash(cells[5]),
        total: cash(cells[9]),
        // Struck through means it has already been deleted on this page.
        deleted: /line-through/.test(tr.getAttribute('style') || ''),
        can_split: !!tr.querySelector('a[id^="split_"]'),
        // The details button turns from yellow to grey once a remittance file
        // is attached. Read as three states, not two: a button with neither
        // class is "don't know", and "don't know" must never read as "done".
        file: (() => {
          const b = tr.querySelector('a[name^="details_"]');
          if (!b) return 'unknown';
          if (b.classList.contains('btn-warning')) return 'missing';
          if (b.classList.contains('btn-default')) return 'attached';
          return 'unknown';
        })(),
      });
    }
    return out;
  }

  /* ── the panel ──────────────────────────────────────────── */
  const ACT = {
    delete: ['#8A1C1C', '#FBE9E9', 'DELETE'],
    split:  ['#6B4E00', '#FFF6DD', 'SPLIT'],
    upload: ['#0B5D1E', '#E4F5E8', 'UPLOAD'],
    done:   ['#6B6B6B', '#F3F3F3', 'has its file'],
    no_file:['#8A1C1C', '#FFF6DD', 'NO FILE — by hand'],
    // Held for you: an unusual in-house ACH, one with a file already on it,
    // or anything left hanging by a split that didn't line up.
    review: ['#6B4E00', '#FFF6DD', 'LEAVE — check it'],
    leave:  ['#3A3A3A', '#FFFFFF', 'leave'],
  };

  /* The panel is built once and kept. Only its body is redrawn, so dragging,
     minimising and scroll position survive a redraw — and there are a lot of
     redraws, because every delete reloads the page. Where it sits is
     remembered in localStorage so it stays put across those reloads. */
  const REMEMBER = 'csPlanBox';
  const recall = () => { try { return JSON.parse(localStorage.getItem(REMEMBER)) || {}; }
                         catch { return {}; } };
  const remember = v => { try { localStorage.setItem(REMEMBER, JSON.stringify(v)); } catch {} };

  function shell() {
    let box = document.getElementById('csPlan');
    if (box) return box;

    const saved = recall();
    box = document.createElement('div');
    box.id = 'csPlan';
    box.style.cssText = `position:fixed;z-index:99999;width:680px;max-width:94vw;
      background:#fff;border:1px solid #999;border-radius:5px;
      box-shadow:0 6px 28px rgba(0,0,0,.3);font:12px/1.45 "Segoe UI",sans-serif;
      display:flex;flex-direction:column;max-height:88vh`;
    // Centred on first use, then wherever it was left.
    box.style.left = (saved.left ?? Math.max(8, (innerWidth - 680) / 2)) + 'px';
    box.style.top  = (saved.top  ?? Math.max(8, innerHeight * 0.06)) + 'px';

    box.innerHTML = `
      <div id="csBar" style="display:flex;align-items:center;gap:8px;cursor:move;
        padding:9px 12px;background:#1F3A5F;color:#fff;border-radius:4px 4px 0 0;
        user-select:none">
        <strong style="flex:1;font-size:13px">The plan for this deposit list</strong>
        <span id="csTick" style="font-size:11px;opacity:.75"></span>
        <button id="csMin" title="Minimise" style="border:0;background:#ffffff26;color:#fff;
          cursor:pointer;padding:2px 9px;border-radius:3px">–</button>
        <button id="csClose" title="Close" style="border:0;background:#ffffff26;color:#fff;
          cursor:pointer;padding:2px 9px;border-radius:3px">×</button>
      </div>
      <div id="csBody" style="overflow:auto;padding:12px 16px"></div>`;
    document.body.appendChild(box);

    const body = box.querySelector('#csBody');
    const setMin = v => {
      body.style.display = v ? 'none' : '';
      box.querySelector('#csMin').textContent = v ? '□' : '–';
      remember({...recall(), min: v});
    };
    setMin(!!saved.min);
    box.querySelector('#csMin').onclick = () => setMin(body.style.display !== 'none');
    box.querySelector('#csClose').onclick = () => { closed = true; box.remove(); };

    // Drag by the title bar, kept inside the window so it can't be lost
    // off-screen after a reload.
    let from = null;
    box.querySelector('#csBar').addEventListener('mousedown', e => {
      if (e.target.tagName === 'BUTTON') return;
      from = {x: e.clientX, y: e.clientY, left: box.offsetLeft, top: box.offsetTop};
      e.preventDefault();
    });
    addEventListener('mousemove', e => {
      if (!from) return;
      const left = Math.min(Math.max(0, from.left + e.clientX - from.x), innerWidth - 120);
      const top  = Math.min(Math.max(0, from.top  + e.clientY - from.y), innerHeight - 40);
      box.style.left = left + 'px';
      box.style.top = top + 'px';
    });
    addEventListener('mouseup', () => {
      if (!from) return;
      from = null;
      remember({...recall(), left: box.offsetLeft, top: box.offsetTop});
    });
    return box;
  }

  function panel(job) {
    if (closed) return;
    const rows = job.deposits || [], plan = job.plan || {}, run = job.run || {};
    const acts = rows.reduce((a, r) => (a[r.action] = (a[r.action] || 0) + 1, a), {});
    const doing = rows.filter(r => r.action !== 'leave');
    const approved = job.approved || {};
    const STATUS = {approved: 'opening…', attached: 'attached — waiting for you',
                    submitted: 'sent, unconfirmed', uploaded: 'sent ✓',
                    failed: 'failed', refused: 'refused'};

    const line = r => {
      const [fg, bg, label] = ACT[r.action] || ACT.leave;
      const file = r.ache && r.ache.file;
      const a = approved[String(r.id)];
      const go = r.action === 'upload'
        ? `<button data-send="${r.id}" style="cursor:pointer;border:1px solid #0B5D1E;
             background:#fff;color:#0B5D1E;border-radius:3px;padding:1px 8px;font-size:11px">${
             a && a.status !== 'approved' ? 'again' : 'Upload'}</button>` : '';
      return `<tr style="background:${bg};color:${fg}">
        <td style="padding:3px 6px;font-family:ui-monospace,monospace">${r.id}</td>
        <td style="padding:3px 6px">${r.type}</td>
        <td style="padding:3px 6px;text-align:right;font-family:ui-monospace,monospace">${usd(r.amount)}</td>
        <td style="padding:3px 6px"><strong>${label}</strong>${
          r.ambiguous ? ' <span style="color:#8A1C1C">⚠ shared amount</span>' : ''}${
          a ? `<div style="font-size:11px;${a.status === 'failed' || a.status === 'refused'
            ? 'color:#8A1C1C' : ''}" title="${(a.message || '').replace(/"/g, '&quot;')}">${
            STATUS[a.status] || a.status}</div>` : ''}</td>
        <td style="padding:3px 6px;font-size:11px">${file ? file : (r.who || '')}</td>
        <td style="padding:3px 6px;text-align:right">${go}</td>
      </tr>`;
    };

    const steps = plan.split_done
      ? [`Upload ${acts.upload || 0} remittance file(s)${
           acts.done ? ` — ${acts.done} of the ${(acts.upload || 0) + acts.done} already have theirs` : ''}`,
         `Delete ${acts.delete || 0} deposit(s) — the in-house parents and any listed wires`]
      : [`Delete the listed wires and the in-house ACHs that aren't being split (${acts.delete || 0})`,
         plan.split_target
           ? `Split deposit ${plan.split_target} into the ${(plan.expected || []).length} ACHs `
             + `(${usd(plan.expected_sum)})`
           : 'Nothing to split — no in-house ACH found',
         'Delete the parent once the children are there',
         'Upload one remittance file to each new deposit'];

    const nothingLeft = !acts.delete && !acts.split && !acts.upload;

    shell().querySelector('#csBody').innerHTML = `
      ${run.stopped ? `<p style="margin:0 0 10px;padding:6px 8px;background:#FBE9E9;
        border-left:3px solid #8A1C1C"><strong>The run stopped.</strong> ${run.stopped}</p>` : ''}

      <p style="margin:0 0 4px"><strong>${plan.split_done
        ? 'The split has already been done.' : 'The split has not been done yet.'}</strong></p>
      <ol style="margin:4px 0 10px;padding-left:20px">${steps.map(t => `<li>${t}</li>`).join('')}</ol>

      ${(plan.problems || []).map(t => `<p style="margin:8px 0;padding:6px 8px;
        background:#FFF6DD;border-left:3px solid #C9A227">${t}</p>`).join('')}

      <p style="margin:8px 0">${acts.review ? `<strong style="color:#6B4E00">${acts.review}
        held for you</strong> · ` : ''}${acts.delete || 0} to delete · ${acts.split || 0} to split ·
        ${acts.upload || 0} to upload · ${acts.done || 0} already uploaded ·${
        acts.no_file ? ` <strong style="color:#8A1C1C">${acts.no_file} with no file</strong> ·` : ''}
        ${acts.leave || 0} left alone</p>

      ${nothingLeft ? `<p style="margin:10px 0;padding:8px 10px;background:#E4F5E8;
          border-left:3px solid #0B5D1E"><strong>Nothing left to do on this list.</strong></p>`
        : `<div style="margin:10px 0;padding:8px 10px;background:#F6F6F6;border:1px solid #ddd;
             border-radius:3px">
        ${run.active
          ? `<strong style="color:#6B4E00">Running…</strong> ${nextLabel(rows, plan)}
             <button id="csStop" style="margin-left:10px;cursor:pointer;border:1px solid #8A1C1C;
               background:#fff;color:#8A1C1C;border-radius:3px;padding:3px 10px">Stop</button>`
          : `<label style="display:block;margin-bottom:8px;cursor:pointer">
               <input type="checkbox" id="csOk" ${okayed ? 'checked' : ''}>
               I've checked the actions above and they're right</label>
             <button id="csGo" ${okayed ? '' : 'disabled'} style="cursor:pointer;border:0;
               border-radius:3px;padding:5px 12px;font-size:12px;${okayed
                 ? 'background:#1F3A5F;color:#fff' : 'background:#ddd;color:#888'}">${
               armed ? 'Sure? click again' : 'Go'}</button>
             <span style="margin-left:10px;color:#555">${nextLabel(rows, plan)}</span>`}
        ${acts.upload ? `<label style="display:block;margin-top:8px;cursor:pointer">
          <input type="checkbox" id="csAuto" ${autoSubmit ? 'checked' : ''}>
          Per-row Upload buttons submit automatically too</label>` : ''}
      </div>`}

      <table style="width:100%;border-collapse:collapse;margin-top:6px">
        <thead><tr style="text-align:left;border-bottom:1px solid #ccc">
          <th style="padding:3px 6px">ID</th><th style="padding:3px 6px">Type</th>
          <th style="padding:3px 6px;text-align:right">Amount</th>
          <th style="padding:3px 6px">Action</th>
          <th style="padding:3px 6px">File / who</th><th></th></tr></thead>
        <tbody>${doing.map(line).join('')}</tbody>
      </table>

      ${(run.log || []).length ? `<details style="margin-top:10px"><summary style="cursor:pointer">
        What the run has done (${run.log.length})</summary>
        <div style="margin-top:6px;font-family:ui-monospace,monospace;font-size:11px">
        ${run.log.slice().reverse().map(l => `${l.action} ${l.id}${
          l.message ? ' — ' + l.message : ''}`).join('<br>')}</div></details>` : ''}

      <details style="margin-top:10px"><summary style="cursor:pointer">
        Why each of these, and the ${rows.length - doing.length} left alone</summary>
        <table style="width:100%;border-collapse:collapse;margin-top:6px"><tbody>
        ${rows.map(r => `<tr><td style="padding:2px 6px;font-family:ui-monospace,monospace;
            vertical-align:top">${r.id}</td>
          <td style="padding:2px 6px">${r.why}</td></tr>`).join('')}</tbody></table></details>

      ${rows.some(r => r.ambiguous) ? `<p style="margin:10px 0;padding:6px 8px;
        background:#FBE9E9;border-left:3px solid #8A1C1C">Some matches land on an amount
        more than one deposit shares, marked ⚠. The figure names a group there, not a row.</p>` : ''}

      <p style="margin:12px 0 0;color:#666">Wires and in-house ACHs are matched on
        amount alone, which is all this list shows. The split children are not: they are
        the ACHs with no split button, paired to files by deposit ID order.</p>`;

    const body = shell().querySelector('#csBody');
    const ok = body.querySelector('#csOk');
    if (ok) ok.onchange = () => {
      okayed = ok.checked;
      // Kept across reloads: every delete reloads this page, and re-ticking
      // the box after each one would make the run unusable.
      remember({...recall(), ok: okayed});
      render();
    };
    const auto = body.querySelector('#csAuto');
    if (auto) auto.onchange = () => { autoSubmit = auto.checked; };
    const stop = body.querySelector('#csStop');
    if (stop) stop.onclick = async () => {
      await ask('/api/upload/run', 'POST', {active: false, reason: 'You stopped it.'});
      run2();
    };
    const go = body.querySelector('#csGo');
    if (go) go.onclick = () => {
      // Two clicks rather than a confirm() dialog: this deletes rows and posts
      // files, and a dialog can cost the click that lets a tab open.
      if (!armed) { armed = true; setTimeout(() => { armed = false; render(); }, 6000); return render(); }
      armed = false;
      // When uploading is what's next, do it straight from this click. A tab
      // opened later, after the page has reloaded a few times, has no click
      // behind it and the browser refuses it.
      const next = nextStep(latest.deposits || [], latest.plan || {});
      if (next.kind === 'upload') { uploadsStarted = true; return sendAll(); }
      start();
    };
    body.querySelectorAll('[data-send]').forEach(btn =>
      btn.onclick = () => send(btn.dataset.send));
  }

  const nextLabel = (rows, plan) => {
    const next = nextStep(rows, plan);
    return next.kind === 'delete' ? `Next: delete deposit ${next.id}.`
         : next.kind === 'split'  ? `Next: split deposit ${next.id} — opens in its own tab.`
         : next.kind === 'upload' ? `Next: upload ${next.count} file(s).`
         : 'Nothing left to do.';
  };

  /* The next thing to do, worked out from the page as it is now rather than
     from a list made earlier. Deletes come first and one at a time, because
     each one reloads the page; then the split; then the uploads. */
  function nextStep(rows, plan) {
    const del = rows.find(r => r.action === 'delete');
    if (del) return {kind: 'delete', id: String(del.id)};
    const sp = rows.find(r => r.action === 'split');
    if (sp) return {kind: 'split', id: String(sp.id)};
    const ups = rows.filter(r => r.action === 'upload');
    if (ups.length) return {kind: 'upload', count: ups.length};
    return {kind: 'done'};
  }

  /* Open the deposit's page and approve it.

     The tab is opened first, synchronously inside the click, because a popup
     opened after an await has lost the click that allowed it and gets
     blocked. Stage 3 waits a moment for the approval to land.

     No "noopener": the deposit page reloads its opener when a file goes in,
     which is what refreshes this list and turns the button grey. A named
     target means a second click reuses the tab instead of piling them up. */
  async function send(id) {
    window.open(`DepositView.cfm?DepositID=${encodeURIComponent(id)}&showSplit=0&viewHistory=0`,
                `cs_deposit_${id}`);
    let r;
    try { r = await ask('/api/upload/approve', 'POST', {ids: [id], auto_submit: autoSubmit}); }
    catch { return alert('The control sheet server did not answer, so nothing was approved.'); }
    if (!r.approved || !r.approved.includes(String(id))) {
      alert(`Deposit ${id} wasn't approved — the plan no longer marks it for upload. `
          + 'The tab that opened will do nothing.');
    }
    run();
  }

  /* Approve every pending upload and open the first one. Stage 3 takes it from
     there, walking the rest in that same tab and closing it at the end.

     Returns the window, or null if the browser refused to open it — which it
     will if this is called from the run rather than from a click. */
  async function sendAll() {
    const ids = (latest.deposits || [])
      .filter(r => r.action === 'upload').map(r => String(r.id));
    if (!ids.length) return null;

    const win = window.open(
      `DepositView.cfm?DepositID=${encodeURIComponent(ids[0])}&showSplit=0&viewHistory=0`,
      'cs_upload');
    let r;
    try { r = await ask('/api/upload/approve', 'POST', {ids, auto_submit: true}); }
    catch {
      alert('The control sheet server did not answer, so nothing was approved.');
      return null;
    }
    if (r.refused && r.refused.length)
      alert(`${r.refused.length} of ${ids.length} weren't approved and will be skipped: `
          + r.refused.join(', '));
    run();
    return win;
  }

  /* ── running the plan ───────────────────────────────────── */
  /* Deleting a row is the page's own delete button plus the confirm it
     raises. The confirm belongs to the page, not to this script, so it has to
     be answered in the page's own world via unsafeWindow — overriding confirm
     here would do nothing at all. Both it and alert are put back afterwards,
     and only for the moment around the click, so a dialog the site raises for
     any other reason still reaches you. */
  function clickDelete(id) {
    const btn = document.getElementById(`deleteRow_${id}`);
    if (!btn) return 'no delete button on that row';
    const w = (typeof unsafeWindow !== 'undefined') ? unsafeWindow : window;
    const realConfirm = w.confirm, realAlert = w.alert;
    const restore = () => { w.confirm = realConfirm; w.alert = realAlert; };
    try {
      w.confirm = () => true;
      w.alert = () => {};
      btn.click();
    } catch (e) {
      restore();
      return e.message;
    }
    setTimeout(restore, 8000);
    addEventListener('beforeunload', restore, {once: true});
    return null;
  }

  async function start() {
    await ask('/api/upload/run', 'POST', {active: true});
    run2();
  }

  /* One action, then stop and wait to be called again.

     Deleting reloads the page and uploading reloads it too, so each pass does
     a single thing and the next page load picks up where this left off. The
     step is recomputed from the list every time, which means a delete that
     didn't take is simply chosen again — and the server counts those, so a
     step that can never succeed stops the run instead of looping. */
  let working = false;
  let splitOpened = false;
  let uploadsStarted = false;
  async function tick(job) {
    if (working || !job.run || !job.run.active) return;
    const next = nextStep(job.deposits || [], job.plan || {});

    if (next.kind === 'done') {
      await ask('/api/upload/run', 'POST', {active: false, reason: null});
      return run2();
    }

    if (next.kind === 'split') {
      // Splitting opens its own tab and doesn't reload this page, so this runs
      // once per page load rather than every time the table redraws.
      if (splitOpened) return;
      splitOpened = true;
      working = true;

      const okay = await ask('/api/upload/split/approve', 'POST',
        {deposit_id: next.id}).catch(() => null);
      if (!okay || !okay.ok) {
        await ask('/api/upload/run', 'POST', {active: false,
          reason: (okay && okay.message) || "Couldn't approve the split."});
        working = false;
        return run2();
      }

      await ask('/api/upload/step', 'POST', {action: 'split', deposit_id: next.id});
      await sleep(PACE);
      const btn = document.getElementById(`split_${next.id}`);
      if (!btn) {
        await ask('/api/upload/run', 'POST', {active: false,
          reason: `Deposit ${next.id} has no split button on this page.`});
        working = false;
        return run2();
      }
      btn.style.outline = '3px solid #C9A227';
      btn.scrollIntoView({block: 'center', behavior: 'smooth'});
      btn.click();
      working = false;
      // The split tab reloads this one when it's done, and the run picks up
      // from there with the parent delete.
      return run2();
    }

    if (next.kind === 'upload') {
      // Carry straight on into the uploads. The tab may be refused, since by
      // now the click that started the run is long gone — in which case say
      // so, and let Go do it, where there is a click.
      if (uploadsStarted) return;
      uploadsStarted = true;
      working = true;
      await sleep(PACE);
      const win = await sendAll();
      await ask('/api/upload/run', 'POST', {active: false, reason: win ? null :
        `Deletes are done and ${next.count} file(s) are ready — press Go to send them. `
        + `The browser wouldn't let a tab open by itself this long after a click.`});
      working = false;
      return run2();
    }

    working = true;
    // A beat between actions. Without it the deletes fire the moment each page
    // settles, which is hard to follow and gives no chance to hit Stop.
    await sleep(PACE);
    const step = await ask('/api/upload/step', 'POST',
      {action: 'delete', deposit_id: next.id}).catch(() => null);
    if (!step || step.stop) { working = false; return run2(); }

    const failed = clickDelete(next.id);
    if (failed) {
      await ask('/api/upload/step', 'POST',
        {action: 'delete', deposit_id: next.id, ok: false, message: failed});
      working = false;
      return run2();
    }
    // The page reloads itself from here. If it doesn't, the observer sees the
    // row go struck-through and comes back round for the next one.
    setTimeout(() => { working = false; run2(); }, 2500);
  }

  /* ── go ─────────────────────────────────────────────────── */
  let closed = false;
  let okayed = !!recall().ok;
  // Off unless ticked: the first runs should stop at the dialog so the file
  // can be looked at before anything is posted.
  let autoSubmit = false;
  // The plan as last returned, so Go knows what's pending, and a
  // re-render can redraw the panel without another round trip.
  let latest = {};
  let armed = false;
  const render = () => latest.deposits && panel(latest);

  async function run() {
    let job;
    try { job = await ask('/api/upload/job'); } catch { return; }
    if (!job || !job.armed) return;

    const rows = readRows();
    if (!rows.length) return;

    let annotated;
    try { annotated = await ask('/api/upload/deposits', 'POST', {deposits: rows}); }
    catch { return; }
    if (!annotated || !annotated.armed) return;

    latest = annotated;
    panel(annotated);
    // Not awaited — but a rejection here would otherwise be silent, and the
    // run would just appear to stall on whatever it last drew.
    tick(annotated).catch(e => console.error('[control sheet] run:', e));
  }

  const run2 = run;
  run().catch(e => console.error('[control sheet] stage 2:', e));

  /* The page redraws the table when a file is uploaded or a row deleted, so
     read it again whenever that happens — otherwise the plan goes on saying
     "upload" for a deposit that has just had its file. Debounced, because a
     DataTable redraw arrives as a burst of mutations, and the panel itself is
     outside the table so redrawing it can't retrigger this. */
  let pending = null;
  const table = document.querySelector('tr[id^="row-"]')?.closest('table');
  if (table) {
    new MutationObserver(() => {
      clearTimeout(pending);
      pending = setTimeout(() => run().catch(
        e => console.error('[control sheet] stage 2:', e)), 400);
    }).observe(table, {subtree: true, childList: true,
                       attributes: true, attributeFilter: ['class', 'style']});
  }
})();