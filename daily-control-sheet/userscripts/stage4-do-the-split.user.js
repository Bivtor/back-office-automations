// ==UserScript==
// @name         Control Sheet — remittance upload, stage 4 (do the split)
// @namespace    control-sheet
// @version      1.0
// @description  On the Split ACH Deposit page, enter the lockbox ACH amounts the control sheet holds and save. Acts only on the deposit the plan approved for splitting.
// @match        https://www.partners.example.com/SimpleAlloc/DailyDeposit/DepositSplitACH.cfm*
// @grant        GM_xmlhttpRequest
// @grant        unsafeWindow
// @connect      127.0.0.1
// @connect      localhost
// @run-at       document-idle
// ==/UserScript==

/*
 * The amounts go in in the order the control sheet holds them, because that is
 * the order the new deposits come out in, and that order is what pairs each
 * new deposit with its remittance file. Two ACHs at the same figure can only
 * be told apart that way.
 *
 * The page's own rule, from its calcRemaining(): the amounts must total
 * strictly less than the deposit, so something is left on the parent. A split
 * that would use it all up is refused here rather than at the alert.
 */

(function () {
  'use strict';

  const SERVER = 'http://127.0.0.1:8777';
  const PACE = 850;
  const sleep = ms => new Promise(r => setTimeout(r, ms));

  const ask = (path, method = 'GET', payload = null) => new Promise((res, rej) => {
    GM_xmlhttpRequest({
      method, url: SERVER + path, timeout: 8000,
      headers: payload ? {'Content-Type': 'application/json'} : {},
      data: payload ? JSON.stringify(payload) : undefined,
      onload: r => { try { res(JSON.parse(r.responseText)); } catch (e) { rej(e); } },
      onerror: rej, ontimeout: rej,
    });
  });

  const cents = n => Math.round(Number(n) * 100);
  const usd = n => '$' + Number(n).toLocaleString('en-US',
    {minimumFractionDigits: 2, maximumFractionDigits: 2});

  let box;
  function say(html, tone = 'work') {
    if (!box) {
      box = document.createElement('div');
      box.style.cssText = `position:fixed;top:12px;right:12px;z-index:99999;max-width:420px;
        font:13px/1.45 "Segoe UI",sans-serif;padding:12px 14px;border-radius:4px;
        box-shadow:0 2px 10px rgba(0,0,0,.2)`;
      document.body.appendChild(box);
    }
    const tones = {work: ['#1F3A5F', '#E8F0FA'], good: ['#0B5D1E', '#E4F5E8'],
                   bad: ['#8A1C1C', '#FBE9E9'], wait: ['#6B4E00', '#FFF6DD']};
    const [fg, bg] = tones[tone] || tones.work;
    Object.assign(box.style, {color: fg, background: bg, border: `1px solid ${fg}33`});
    box.innerHTML = html;
  }

  const tell = (status, message) =>
    ask('/api/upload/split/report', 'POST', {status, message}).catch(() => {});

  const $id = i => document.getElementById(i);
  const rowCount = () => Number($id('SplitACHCount')?.value || 0);

  /* Rows are added by the page's own + button, which is what keeps
     SplitACHCount and the field names in step. Building the markup here
     instead would leave its counter behind and the post would drop rows. */
  async function makeRows(n) {
    const add = $id('AddSplitACHBtn');
    let guard = 0;
    while (rowCount() < n && guard++ < 200) {
      add.click();
      await sleep(40);
    }
    return rowCount() >= n;
  }

  async function fill(amounts) {
    for (let i = 0; i < amounts.length; i++) {
      const el = $id(`SplitACH_Amount_${i + 1}`);
      if (!el) return `row ${i + 1} never appeared`;
      // Plain digits: the page multiplies the raw value by 100, so a comma
      // would make it NaN and quietly wreck the remaining-amount sum.
      el.value = Number(amounts[i]).toFixed(2);
      el.dispatchEvent(new Event('input', {bubbles: true}));
      el.dispatchEvent(new Event('change', {bubbles: true}));
      el.dispatchEvent(new Event('blur', {bubbles: true}));
      await sleep(30);
    }
    return null;
  }

  async function run() {
    const depositId = $id('DepositID')?.value
      || new URLSearchParams(location.search).get('DepositID') || '';
    if (!depositId) return;

    let job;
    for (let n = 0; n < 8; n++) {
      try { job = await ask('/api/upload/job'); } catch { return; }
      if (!job.armed) return;
      if (job.split && job.split.deposit_id) break;
      await sleep(750);
    }
    const s = job.split || {};
    if (String(s.deposit_id) !== String(depositId)) return;   // not ours
    if (s.status && s.status !== 'approved') {
      return say(`This split was already ${s.status}. Nothing done.`, 'wait');
    }

    const amounts = (job.aches || []).map(a => a.amount);
    const initAmt = Number($id('initAmt')?.value || 0);
    const total = amounts.reduce((n, a) => n + Number(a || 0), 0);

    const wrong = [];
    if (!amounts.length) wrong.push('The control sheet has no ACHs to split out.');
    if (!$id('AddSplitACHBtn') || !$id('SplitACH_Amount_1'))
      wrong.push("This page doesn't have the split form on it.");
    if (!initAmt) wrong.push("Couldn't read the deposit's amount off the page.");
    else if (cents(total) >= cents(initAmt))
      wrong.push(`The ${amounts.length} ACHs total ${usd(total)} but this deposit is `
               + `${usd(initAmt)}. The page needs something left over, so it would refuse this.`);
    // A form that already has rows or a filled first box has been worked on;
    // adding to it would double the split.
    if (rowCount() !== 1 || ($id('SplitACH_Amount_1')?.value || '').trim() !== '')
      wrong.push('This split form has already been started — not adding to it.');

    if (wrong.length) {
      await tell('refused', wrong.join(' '));
      return say(`<strong>Not splitting anything.</strong><ul style="margin:6px 0 0;padding-left:18px">
        ${wrong.map(t => `<li>${t}</li>`).join('')}</ul>`, 'bad');
    }

    say(`Entering ${amounts.length} amounts totalling <strong>${usd(total)}</strong>,
      leaving <strong>${usd(initAmt - total)}</strong> on deposit ${depositId}…`, 'wait');
    await sleep(PACE);

    if (!await makeRows(amounts.length)) {
      await tell('failed', 'Could not add enough rows.');
      return say('Could not add enough rows to the form. Nothing saved.', 'bad');
    }
    const bad = await fill(amounts);
    if (bad) {
      await tell('failed', bad);
      return say(`Filling the amounts failed: ${bad}. Nothing saved.`, 'bad');
    }

    // Read back what the page makes of it, and require it to agree with our
    // own arithmetic to the cent.
    //
    // Not just "is it positive": calcRemaining leaves remainingAmount at its
    // last good value when it rejects, so a stale positive number can sit
    // there while the real remainder is negative. Agreement is the test that
    // actually proves every amount went into the row it was meant to.
    await sleep(PACE);
    const remaining = Number($id('remainingAmount')?.value || 0);
    const expected = initAmt - total;
    if (Math.abs(cents(remaining) - cents(expected)) > 1) {
      await tell('refused',
        `The page has ${remaining} remaining, we make it ${expected.toFixed(2)}.`);
      return say(`The page says <strong>${usd(remaining)}</strong> would be left, but the
        amounts we entered make it <strong>${usd(expected)}</strong>. Something didn't go
        in as intended, so nothing was saved.`, 'bad');
    }

    say(`${amounts.length} amounts in, ${usd(remaining)} left on the parent. Saving…`, 'wait');
    await tell('filled', `Remaining ${remaining}`);
    await sleep(PACE);

    // The page's own Save button, so its validation runs and it submits the
    // form the way it expects.
    const save = [...document.querySelectorAll('button')]
      .find(b => /save/i.test(b.textContent) && /calcRemaining/.test(b.getAttribute('onclick') || ''));
    if (!save) {
      await tell('failed', 'No Save button found.');
      return say('Filled in, but the Save button is missing — press it yourself.', 'bad');
    }

    // The alert only fires when the page rejects the amounts; catching it
    // turns a silent refusal into something recorded.
    const w = (typeof unsafeWindow !== 'undefined') ? unsafeWindow : window;
    const realAlert = w.alert;
    let complaint = null;
    w.alert = m => { complaint = String(m); };
    save.click();
    setTimeout(() => { w.alert = realAlert; }, 5000);

    await sleep(600);
    if (complaint) {
      await tell('failed', complaint);
      return say(`The page refused it: <strong>${complaint}</strong> Nothing saved.`, 'bad');
    }

    await tell('submitted', `Split into ${amounts.length}, ${remaining} left.`);
    say(`Split saved — ${amounts.length} new deposits, ${usd(remaining)} left on
      ${depositId}. Sending you back to the list…`, 'good');

    // The list tab won't know about the new deposits until it reloads, and it
    // is the tab that carries the rest of the run.
    setTimeout(() => {
      try {
        if (window.opener && !window.opener.closed) {
          window.opener.location.reload();
          window.opener.focus();
        }
      } catch {}
    }, 2500);
  }

  // Without this a throw is a silent rejection, and the only sign of it is
  // whichever banner was last on screen — which looks like a hang.
  run().catch(e => {
    say(`<strong>This stopped with an error:</strong> ${e && e.message ? e.message : e}
      <div style="margin-top:6px">Nothing was saved. Do the split by hand.</div>`, 'bad');
    console.error('[control sheet] stage 4:', e);
  });
})();