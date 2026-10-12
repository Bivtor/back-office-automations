// ==UserScript==
// @name         Control Sheet — remittance upload, stage 1 (find the session)
// @namespace    control-sheet
// @version      1.0
// @description  On the Daily Deposit Summary, find today's ACH report upload and open its details page. Does nothing unless the control sheet has armed a job.
// @match        https://www.partners.example.com/SimpleAlloc/DailyDeposit/DailyDepositSummary.cfm*
// @grant        GM_xmlhttpRequest
// @connect      127.0.0.1
// @connect      localhost
// @run-at       document-idle
// ==/UserScript==

/*
 * There is no marker in the URL and none is wanted. The script asks the local
 * control sheet server whether a job is armed; no server, or no job, and it
 * stops without touching the page. That makes the page safe to open by hand on
 * any other day, and means nothing has to be remembered about how it was
 * opened.
 *
 * GM_xmlhttpRequest rather than fetch, for two reasons. It sidesteps CORS, so
 * the server needs no permissive headers — which matters, because a localhost
 * server sending Access-Control-Allow-Origin:* would let any site you happen to
 * have open read the remittance files. It also avoids Private Network Access:
 * this page is on a private address reaching localhost, and Chromium wants
 * Access-Control-Allow-Private-Network on the preflight for that hop.
 */

(function () {
  'use strict';

  const SERVER = 'http://127.0.0.1:8777';

  const ask = (path, method = 'GET', payload = null) => new Promise((res, rej) => {
    GM_xmlhttpRequest({
      method, url: SERVER + path, timeout: 4000,
      headers: payload ? {'Content-Type': 'application/json'} : {},
      data: payload ? JSON.stringify(payload) : undefined,
      onload: r => { try { res(JSON.parse(r.responseText)); } catch (e) { rej(e); } },
      onerror: rej, ontimeout: rej,
    });
  });

  /* ── the banner ─────────────────────────────────────────── */
  let box;
  function say(html, tone = 'work') {
    if (!box) {
      box = document.createElement('div');
      box.style.cssText = `position:fixed;top:12px;right:12px;z-index:99999;max-width:380px;
        font:13px/1.45 "Segoe UI",sans-serif;padding:12px 14px;border-radius:4px;
        box-shadow:0 2px 10px rgba(0,0,0,.2);white-space:normal`;
      document.body.appendChild(box);
    }
    const tones = {work: ['#1F3A5F', '#E8F0FA'], good: ['#0B5D1E', '#E4F5E8'],
                   bad: ['#8A1C1C', '#FBE9E9']};
    const [fg, bg] = tones[tone] || tones.work;
    box.style.color = fg;
    box.style.background = bg;
    box.style.border = `1px solid ${fg}33`;
    box.innerHTML = html;
  }

  /* ── matching the row ───────────────────────────────────── */
  // "25  ACHs Afternoon Report.xls" arrives with a double space, and the report
  // is sometimes the morning pull. Compare on collapsed whitespace, ignoring
  // case, so neither costs us the match.
  const norm = s => String(s || '').replace(/\s+/g, ' ').trim().toLowerCase();

  const todayUS = () => {
    const d = new Date();
    return `${String(d.getMonth() + 1).padStart(2, '0')}/${
      String(d.getDate()).padStart(2, '0')}/${d.getFullYear()}`;
  };

  /* Rows whose filename looks like the afternoon report and which carry
     today's date.

     Matched on "achs afternoon report" appearing in the name rather than on
     the whole filename, because the leading item count changes daily ("25  ACHs
     Afternoon Report.xls", "41  ACHs Afternoon Report.xls") and the sheet may
     have been built from the morning pull — but the row to find here is always
     the afternoon one.

     The name is matched first and the date only narrows it, since the name is
     the specific thing and today's date is shared by everything uploaded
     today. A row is only taken if it is the single survivor — see below. */
  function candidates(pattern) {
    const want = norm(pattern), today = todayUS();
    const out = [];
    for (const tr of document.querySelectorAll('tr[id^="row-"]')) {
      const cells = [...tr.cells].map(td => td.textContent);
      const titles = [...tr.querySelectorAll('[title]')].map(el => el.title);
      const named = [...cells, ...titles].some(c => norm(c).includes(want));
      if (!named) continue;
      const dated = cells.some(c => c.includes(today));
      const link = tr.querySelector('a[href*="DailyDepositDetails.cfm?UploadSessionID="]');
      if (link) out.push({tr, link, dated, id: tr.id.replace('row-', '')});
    }
    // Prefer the ones dated today; fall back to name-only so a date column
    // that reads differently than expected is visible rather than silent.
    const dated = out.filter(r => r.dated);
    return {all: out, dated, use: dated.length ? dated : out};
  }

  const row = r => `<div style="margin-top:6px;font-family:ui-monospace,monospace;font-size:12px">
    session ${r.id}${r.dated ? '' : ' (not dated today)'}</div>`;

  /* ── go ─────────────────────────────────────────────────── */
  (async function run() {
    let job;
    try {
      job = await ask('/api/upload/job');
    } catch {
      return;   // control sheet isn't running: this is just a normal visit
    }
    if (!job || !job.armed) return;

    const pattern = job.report_pattern || 'achs afternoon report';
    say(`Looking for the <strong>${pattern}</strong> uploaded ${todayUS()}…`);

    const {all, dated, use} = candidates(pattern);

    if (!use.length) {
      return say(`No row on this page looks like the <strong>${pattern}</strong>. `
        + 'If the afternoon report hasn\'t been uploaded yet, upload it and reload. '
        + 'Nothing was clicked.', 'bad');
    }

    if (use.length > 1) {
      return say(`<strong>${use.length} rows</strong> match "${pattern}"`
        + `${dated.length > 1 ? ' and today\'s date' : ''}. Not guessing which — `
        + 'open the right one yourself.' + use.map(row).join(''), 'bad');
    }

    const pick = use[0];
    pick.tr.style.outline = '2px solid #0B5D1E';
    pick.tr.scrollIntoView({block: 'center', behavior: 'smooth'});

    // Record which session we're working, so stage 2 can confirm it landed on
    // the same one rather than trusting whatever page it wakes up on.
    try { await ask('/api/upload/session', 'POST', {session_id: pick.id}); }
    catch { /* stage 2 will re-check anyway */ }

    const warn = dated.length ? '' :
      '<br><em>Careful: this row isn\'t dated today.</em>';
    say(`Found session <strong>${pick.id}</strong>`
      + ` — ${job.files.length} file${job.files.length === 1 ? '' : 's'} to upload.`
      + `${warn}<br>Opening its details page…`, dated.length ? 'good' : 'work');

    setTimeout(() => { pick.link.click(); }, dated.length ? 900 : 2500);
  })();
})();