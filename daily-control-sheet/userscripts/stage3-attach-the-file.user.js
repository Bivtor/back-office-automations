// ==UserScript==
// @name         Control Sheet — remittance upload, stage 3 (attach the file)
// @namespace    control-sheet
// @version      1.0
// @description  On a deposit's DepositView page, attach the remittance file the control sheet approved for it. Acts only on deposits approved from the plan panel; by default attaches and waits for you to press Upload.
// @match        https://www.partners.example.com/SimpleAlloc/DailyDeposit/DepositView.cfm*
// @grant        GM_xmlhttpRequest
// @connect      127.0.0.1
// @connect      localhost
// @run-at       document-idle
// ==/UserScript==

/*
 * How the file gets there without the picker: a path can never be assigned to
 * a file input, but a FileList built from bytes can. The bytes come from the
 * control sheet server, get wrapped in a File with the real name and MIME
 * type, and are loaded into the form's own input. The form then posts exactly
 * as it would after a manual pick — a multipart body carries a field name, a
 * filename and bytes, and nothing records where they came from.
 *
 * What this checks before touching anything, because the destination is baked
 * into the page and there is no undo for a file posted into the wrong folder:
 *   - the deposit was approved from the plan panel, and not already sent
 *   - the form's DepositID is this page's deposit
 *   - its FolderName is that ID zero-padded to eight digits
 *   - the page's Total Amount is the amount the file was matched on
 *   - the deposit hasn't been deleted
 * Any mismatch stops it with the reason on screen.
 */

(function () {
  'use strict';

  const SERVER = 'http://127.0.0.1:8777';

  const ask = (path, method = 'GET', payload = null, responseType = 'text') =>
    new Promise((res, rej) => {
      GM_xmlhttpRequest({
        method, url: SERVER + path, timeout: 20000, responseType,
        headers: payload ? {'Content-Type': 'application/json'} : {},
        data: payload ? JSON.stringify(payload) : undefined,
        onload: r => {
          if (r.status !== 200) return rej(new Error(`server said ${r.status}`));
          if (responseType === 'arraybuffer') return res(r.response);
          try { res(JSON.parse(r.responseText)); } catch (e) { rej(e); }
        },
        onerror: rej, ontimeout: () => rej(new Error('timed out')),
      });
    });

  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const cents = n => n == null ? null : Math.round(Number(n) * 100);
  const usd = n => '$' + Number(n).toLocaleString('en-US',
    {minimumFractionDigits: 2, maximumFractionDigits: 2});

  /* ── banner ─────────────────────────────────────────────── */
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

  const tell = (id, status, message) =>
    ask('/api/upload/report', 'POST', {deposit_id: id, status, message}).catch(() => {});

  /* ── reading the page ───────────────────────────────────── */
  function pageFacts() {
    const form = document.getElementById('UploadForm');
    const val = n => form?.querySelector(`[name="${n}"]`)?.value?.trim() || '';
    const heading = document.querySelector('.panel-heading')?.textContent || '';
    const total = heading.match(/Total Amount:\s*\$?([\d,]+\.\d{2})/i);

    // "Deleted:" sits in a label column with the detail in the next one; a
    // live deposit leaves that detail empty.
    let deleted = '';
    for (const b of document.querySelectorAll('.panel-body b')) {
      if (/^\s*Deleted:\s*$/i.test(b.textContent)) {
        deleted = b.closest('.row')?.querySelector('.col-md-6')?.textContent.trim() || '';
      }
    }
    return {
      form,
      input: form?.querySelector('input[type="file"][name="RemittanceFile"]'),
      submit: form?.querySelector('input[type="submit"]'),
      depositId: val('DepositID'),
      folder: val('FolderName'),
      urlId: new URLSearchParams(location.search).get('DepositID') || '',
      total: total ? parseFloat(total[1].replace(/,/g, '')) : null,
      deleted,
    };
  }

  function problems(p, job) {
    const out = [];
    if (!p.form || !p.input || !p.submit)
      out.push("This page has no upload form, so there's nothing to attach to.");
    if (p.depositId !== p.urlId)
      out.push(`The form uploads to deposit ${p.depositId || '(none)'}, but this page is ${p.urlId}.`);
    const wantFolder = p.depositId.padStart(8, '0');
    if (p.folder !== wantFolder)
      out.push(`The upload folder is ${p.folder || '(none)'}, expected ${wantFolder}.`);
    if (p.total == null)
      out.push("Couldn't read the deposit's Total Amount off the page.");
    else if (cents(p.total) !== cents(job.expected_amount))
      out.push(`This deposit is ${usd(p.total)}, but the file was matched to ${usd(job.expected_amount)}.`);
    if (p.deleted)
      out.push(`This deposit has been deleted (${p.deleted}).`);
    return out;
  }

  /* Files already on this deposit, from the page's own Remittance table.

     This is the confirmation. The site reloads the whole page after an
     upload, so the attached file shows up here on the next load — which is
     the site's own record of what it holds, rather than anything inferred
     from a response. Rows are the file name, the upload date and who did it. */
  const tidy = t => String(t || '').replace(/\s+/g, ' ').trim();
  function attachedFiles() {
    return [...document.querySelectorAll('#tblAttachments tbody tr')]
      .map(tr => [...tr.cells].map(td => tidy(td.textContent)))
      .filter(c => c.length >= 3 && c[0] && !/no data/i.test(c[0]))
      .map(c => ({name: c[0], date: c[1], by: c[2]}));
  }
  const findAttached = name =>
    attachedFiles().find(f => f.name.toLowerCase() === tidy(name).toLowerCase());

  /* ── the upload ─────────────────────────────────────────── */
  async function attach(p, job) {
    const f = job.file;
    const bytes = await ask(f.url, 'GET', null, 'arraybuffer');
    if (!bytes || bytes.byteLength !== f.size)
      throw new Error(`got ${bytes ? bytes.byteLength : 0} bytes, expected ${f.size}`);

    // Name, type and bytes are the three things a multipart upload carries.
    // The type is passed explicitly: left out it would be "", which a server
    // checking the upload's type can reject.
    const file = new File([bytes], f.name, {type: f.type, lastModified: Date.now()});
    const dt = new DataTransfer();
    dt.items.add(file);
    p.input.files = dt.files;
    p.input.dispatchEvent(new Event('input', {bubbles: true}));
    p.input.dispatchEvent(new Event('change', {bubbles: true}));

    if (p.input.files.length !== 1 || p.input.files[0].name !== f.name)
      throw new Error("the file didn't stay attached to the input");
  }

  /* The Remittance table still decides whether a file landed. But when it
     didn't, the only account of what happened is what the site said in the
     response frame, so that gets read and recorded on the way past. Without it
     a failure is just "not there", with no way to tell a rejected upload from
     one that was never really sent. */
  function watchResponse(id, fileName) {
    const frame = document.querySelector('iframe[name="uploadFrame"]');
    if (!frame) return;
    frame.addEventListener('load', () => {
      let text = '';
      try { text = frame.contentDocument?.body?.innerText?.trim() || '(empty response)'; }
      catch { text = '(response not readable)'; }
      tell(id, 'submitted', `Sent ${fileName}. The site replied: ${text.slice(0, 300)}`);
      setTimeout(() => location.reload(), 1500);
    }, {once: true});
  }

  /* Move to the next deposit in the queue, or finish.

     Navigating this same tab rather than opening a new one per deposit: a tab
     opened without a click behind it is blocked, and after the first upload
     there is no click left. Setting location on a tab we are already in needs
     no such permission. The server holds the place in the queue, so the page
     reload the site does after every upload can't lose it. */
  async function nextOrFinish(currentId, fileName) {
    let q;
    try { q = await ask(`/api/upload/next?after=${encodeURIComponent(currentId)}`); }
    catch { return; }

    if (q && q.deposit_id) {
      say(`<strong>✓ ${fileName}</strong> done. ${q.left} left —
        opening deposit <strong>${q.deposit_id}</strong>…`, 'good');
      setTimeout(() => {
        location.href = `DepositView.cfm?DepositID=${encodeURIComponent(q.deposit_id)}`
                      + `&showSplit=0&viewHistory=0`;
      }, 1800);
      return;
    }

    say(`<strong>✓ ${fileName}</strong> done, and nothing left in the queue.
      Closing this tab…`, 'good');
    setTimeout(() => {
      window.close();
      // close() is refused if this tab wasn't opened by a script, so say so
      // rather than sitting on a banner that claims it's closing.
      setTimeout(() => say(`<strong>All uploads done.</strong> This tab couldn't
        close itself — you can close it.`, 'good'), 800);
    }, 2500);
  }

  function openDialog() {
    // The page's own button, so Bootstrap opens the modal the way it expects.
    document.querySelector('[data-target="#uploadModal"]')?.click();
  }

  /* ── go ─────────────────────────────────────────────────── */
  async function run() {
    const p = pageFacts();
    if (!p.urlId) return;

    // The plan panel opens this tab before telling the server, so the
    // approval can lag the page by a moment. Look a few times before
    // concluding this deposit wasn't approved.
    let job;
    for (let n = 0; n < 8; n++) {
      try { job = await ask(`/api/upload/for/${encodeURIComponent(p.urlId)}`); }
      catch { return; }                     // no server: an ordinary visit
      if (!job.armed) return;
      if (job.approved) break;
      await sleep(750);
    }
    if (!job || !job.approved) return;     // not one of ours: leave the page alone

    // The table outranks any status we recorded: if the file is on the
    // deposit it's done, however we got here, and it is never sent twice.
    const there = findAttached(job.file.name);
    if (there) {
      if (job.status !== 'uploaded') await tell(p.urlId, 'uploaded',
        `On the deposit: uploaded ${there.date} by ${there.by}.`);
      say(`<strong>✓ ${job.file.name}</strong> is attached to deposit
        <strong>${p.urlId}</strong> — uploaded ${there.date} by ${there.by}.`, 'good');
      if (job.auto_submit) await nextOrFinish(p.urlId, job.file.name);
      return;
    }

    // Submitted but not in the table means the post went out and the file
    // didn't land. That's worth saying plainly rather than hedging.
    if (job.status === 'submitted' || job.status === 'uploaded') {
      const said = job.message ? `<div style="margin-top:6px;font-size:12px">${job.message}</div>` : '';
      await tell(p.urlId, 'failed',
        `Not on the deposit after submitting. ${job.message || '(no reply recorded)'}`);
      say(`<strong>${job.file.name}</strong> was submitted to deposit
        <strong>${p.urlId}</strong> but <strong>isn't in its Remittance list.</strong>
        The upload didn't land.${said}
        <div style="margin-top:6px">Approve it again from the plan to retry.</div>`, 'bad');
      // The rest of the queue still runs. One deposit that wouldn't take its
      // file shouldn't cost the others, and this one stays red in the plan.
      if (job.auto_submit) setTimeout(() => nextOrFinish(p.urlId, job.file.name), 5000);
      return;
    }

    const wrong = problems(p, job);
    if (wrong.length) {
      await tell(p.urlId, 'refused', wrong.join(' '));
      say(`<strong>Not attaching anything.</strong><ul style="margin:6px 0 0;padding-left:18px">
        ${wrong.map(t => `<li>${t}</li>`).join('')}</ul>`, 'bad');
      // Recorded as refused, so the plan shows it in red. The rest of the
      // queue still runs — one bad deposit shouldn't strand the others.
      if (job.auto_submit) setTimeout(() => nextOrFinish(p.urlId, job.file.name), 4000);
      return;
    }

    say(`Fetching <strong>${job.file.name}</strong>…`);
    try {
      await attach(p, job);
    } catch (e) {
      await tell(p.urlId, 'failed', `Couldn't attach: ${e.message}`);
      say(`<strong>Couldn't attach the file:</strong> ${e.message}`, 'bad');
      if (job.auto_submit) setTimeout(() => nextOrFinish(p.urlId, job.file.name), 4000);
      return;
    }
    await tell(p.urlId, 'attached', null);

    // Recorded at the moment of submitting, because the page is about to be
    // replaced and nothing after this point is guaranteed to run.
    let submitted = false;
    p.form.addEventListener('submit', () => {
      submitted = true;
      tell(p.urlId, 'submitted', `Sent ${job.file.name}.`);
      watchResponse(p.urlId, job.file.name);
      say(`Uploading <strong>${job.file.name}</strong> to deposit
        <strong>${p.urlId}</strong>… the page will reload and confirm it.`, 'wait');
    }, {once: true});

    if (!job.auto_submit) {
      // The dialog is only opened when you're the one pressing Upload. On an
      // automatic run it is left shut: this page loads bootstrap-filestyle,
      // which rebuilds file inputs, and showing the modal is a chance for it
      // to rebuild the one we just put a file on. The form posts perfectly
      // well without ever being seen.
      openDialog();
      return say(`Attached <strong>${job.file.name}</strong>
        (${(job.file.size / 1024).toFixed(1)} KB) to deposit <strong>${p.urlId}</strong>
        at ${usd(p.total)}.<br><br>Check it in the dialog, then press
        <strong>Upload</strong>.`, 'good');
    }

    say(`Attached <strong>${job.file.name}</strong>. Submitting…`, 'wait');
    await sleep(500);

    // The file has to still be on the input at the instant of submitting.
    // An empty input posts a form that looks fine and lands nothing, which is
    // exactly what "submitted but not in the Remittance list" looks like.
    if (p.input.files.length !== 1 || p.input.files[0].name !== job.file.name) {
      say('The file came off the input — putting it back…', 'wait');
      try { await attach(p, job); }
      catch (e) {
        await tell(p.urlId, 'failed', `File wouldn't stay on the input: ${e.message}`);
        say(`<strong>The file wouldn't stay attached</strong> (${e.message}),
          so nothing was submitted.`, 'bad');
        return nextOrFinish(p.urlId, job.file.name);
      }
    }

    // requestSubmit with the button, not form.submit(): it fires the form's
    // submit handlers and sends the button's own name=value ("submit=Upload")
    // with the post, both of which a plain submit() skips. (form.submit is
    // shadowed here anyway — the button is named "submit".)
    try {
      if (p.form.requestSubmit) p.form.requestSubmit(p.submit);
      else p.submit.click();
    } catch { p.submit.click(); }

    // If the submit event never fired, the post never happened — say that,
    // rather than leaving it recorded as sent and blaming the site later.
    await sleep(1200);
    if (!submitted) {
      p.submit.click();                       // once more, as a real click
      await sleep(1200);
    }
    if (!submitted) {
      await tell(p.urlId, 'failed',
        'The form never submitted — neither requestSubmit nor a click on Upload fired it.');
      say(`<strong>The form never submitted.</strong> Nothing was sent to deposit
        ${p.urlId}. Press Upload yourself, or approve it again from the plan.`, 'bad');
      return nextOrFinish(p.urlId, job.file.name);
    }
  }

  /* Anything that throws in there would otherwise be a silent rejection, and
     the only sign of it would be whichever banner happened to be showing —
     which reads as a hang. Say what broke instead. */
  run().catch(e => {
    say(`<strong>This stopped with an error:</strong> ${e && e.message ? e.message : e}
      <div style="margin-top:6px">Nothing further was submitted. Press Upload yourself,
      or approve it again from the plan.</div>`, 'bad');
    console.error('[control sheet] stage 3:', e);
  });
})();