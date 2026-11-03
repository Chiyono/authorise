// Invoice authorisation chase e-mail generator.
//
// Drop the COINS export (.xlsx) of invoices awaiting authorisation, press Generate, and an Outlook-ready
// e-mail is built, shown on the page and copied to the clipboard (as formatted HTML, so the table keeps its
// look when pasted into a new Outlook message).  The subject line is copied with its own small button.
//
// Expected columns (found by their heading, so their order doesn't matter):
//   Contract / Dept | PO Num | Name | Supplier Reference | Invoice Date | Due Date | Balance
//
// The wording follows the two example e-mails.  To change the wording, edit buildEmail() below.
(function () {
  'use strict';

  const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const DAY_MS = 86400000;
  const NUMBER_WORDS = ['no', 'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine', 'ten', 'eleven', 'twelve'];

  // ---- dates (all kept as UTC midnight so time zones can't move them a day) ----
  const utc = (y, m, d) => new Date(Date.UTC(y, m, d));
  const addDays = (d, n) => new Date(d.getTime() + n * DAY_MS);
  const ordinal = n => n + (n % 100 >= 11 && n % 100 <= 13 ? 'th' : ({ 1: 'st', 2: 'nd', 3: 'rd' }[n % 10] || 'th'));
  const longDay = d => DAYS[d.getUTCDay()] + ' ' + ordinal(d.getUTCDate()) + ' ' + MONTHS[d.getUTCMonth()];   // Thursday 12th November
  const dayMonth = d => ordinal(d.getUTCDate()) + ' ' + MONTHS[d.getUTCMonth()];                          // 19th November
  const shortDate = d => d.getUTCDate() + ' ' + MONTHS[d.getUTCMonth()].slice(0, 3);                       // 19 Nov

  // Spreadsheet cell -> date.  Excel stores dates as day numbers (46314 = 19 Oct 2026), but text and real dates are handled too.
  function toDate(v) {
    if (v instanceof Date && !isNaN(v)) return utc(v.getFullYear(), v.getMonth(), v.getDate());
    if (typeof v === 'number' && isFinite(v)) return new Date(Date.UTC(1899, 11, 30) + Math.round(v) * DAY_MS);
    if (typeof v === 'string') {
      const s = v.trim();
      let m = /^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})$/.exec(s);
      if (m) return utc(m[3].length === 2 ? 2000 + +m[3] : +m[3], +m[2] - 1, +m[1]);
      m = /^(\d{4})-(\d\d)-(\d\d)/.exec(s);
      if (m) return utc(+m[1], +m[2] - 1, +m[3]);
      if (/^\d+(\.\d+)?$/.test(s)) return toDate(parseFloat(s));
    }
    return null;
  }

  const money = n => '£' + n.toLocaleString('en-GB', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  // ---- spreadsheet rows -> invoices ----
  // rows: array of arrays (first non-empty row = headings)
  function parseRows(rows) {
    const norm = s => String(s == null ? '' : s).trim().toLowerCase();
    const hi = rows.findIndex(r => r && r.some(c => norm(c) === 'supplier reference') && r.some(c => norm(c) === 'balance'));
    if (hi < 0) throw new Error('I can\'t find the column headings (Supplier Reference, Balance...) - is this the usual export?');
    const head = rows[hi].map(norm);
    const col = (...names) => { for (const n of names) { const i = head.indexOf(n); if (i >= 0) return i; } return -1; };
    const c = {
      contract: col('contract / dept', 'contract'), po: col('po num', 'po number'), name: col('name'),
      ref: col('supplier reference'), inv: col('invoice date'), due: col('due date'), bal: col('balance')
    };
    const missing = Object.keys(c).filter(k => c[k] < 0);
    if (missing.length) throw new Error('Missing column(s): ' + missing.join(', ') + '.');

    const out = [];
    rows.slice(hi + 1).forEach(r => {
      if (!r) return;
      const bal = typeof r[c.bal] === 'number' ? r[c.bal] : parseFloat(String(r[c.bal] == null ? '' : r[c.bal]).replace(/[£,\s]/g, ''));
      const name = String(r[c.name] == null ? '' : r[c.name]).trim();
      if (!name || !isFinite(bal)) return;
      out.push({
        contract: String(r[c.contract] == null ? '' : r[c.contract]).trim(),
        po: String(r[c.po] == null ? '' : r[c.po]).trim(),
        supplier: name,
        ref: String(r[c.ref] == null ? '' : r[c.ref]).replace(/−/g, '–').trim(),     // COINS uses a minus sign; the e-mail uses an en dash
        invDate: toDate(r[c.inv]), due: toDate(r[c.due]), balance: bal
      });
    });
    out.forEach(i => { if (/Park Electrical/i.test(i.supplier) && i.invDate) i.due = eomDue(i.invDate); });
    if (!out.length) throw new Error('No invoices found in that spreadsheet.');
    return out;
  }

  // Park Electrical's terms are always 30 days end of month, whatever COINS shows as the due date
  const eomDue = d => addDays(utc(d.getUTCFullYear(), d.getUTCMonth() + 1, 0), 30);

  // "Access Platforms (Aberdeen) Ltd" -> "Access Platforms"
  const shortName = s => s.replace(/\s*\([^)]*\)/g, '').replace(/\s+(Ltd|Limited|PLC|LLP)\.?$/i, '').replace(/\s+UK$/i, '').trim();

  const joinList = (items, last) => items.length < 2 ? items.join('') : items.slice(0, -1).join(', ') + ' ' + last + ' ' + items[items.length - 1];

  // The next Monday after the given day (what the page proposes as the payment run date)
  function nextMonday(from) {
    const d = utc(from.getFullYear(), from.getMonth(), from.getDate());
    return addDays(d, ((8 - d.getUTCDay()) % 7) || 7);
  }

  // ---- the e-mail ----
  // opts.runDate  the payment run day (UTC date).  It captures everything due up to 3 days later.
  // opts.today    today (UTC date), used to avoid asking for a reply by a day that has gone
  function buildEmail(invoices, opts) {
    const runDate = opts.runDate;
    const today = opts.today;
    const cutoff = addDays(runDate, 3);          // Monday's run captures everything due up to Thursday
    const replyBy = addDays(runDate, -3);        // the Friday before
    const runDay = DAYS[runDate.getUTCDay()];

    const list = invoices.slice().sort((a, b) => (a.due ? a.due.getTime() : Infinity) - (b.due ? b.due.getTime() : Infinity));
    const contracts = [];
    list.forEach(i => { if (i.contract && !contracts.includes(i.contract)) contracts.push(i.contract); });
    const multi = contracts.length > 1;
    const total = list.reduce((a, i) => a + i.balance, 0);
    // overdue: due before today (top priority).  soon: not overdue yet, but due on or before the run's cut-off,
    // so it must make this payment run or it will fall overdue.
    const overdue = list.filter(i => i.due && i.due.getTime() < today.getTime());
    const soon = list.filter(i => i.due && i.due.getTime() >= today.getTime() && i.due.getTime() <= cutoff.getTime());
    const inside = overdue.concat(soon);
    const isOverdue = i => overdue.includes(i);
    const isInside = i => inside.includes(i);
    const daysLate = i => Math.round((today.getTime() - i.due.getTime()) / DAY_MS);
    const describe = (i, late) => esc(shortName(i.supplier)) + ' (' + (multi && i.contract ? 'Contract ' + esc(i.contract) + ', ' : '') + money(i.balance) +
      ', ' + (late ? 'was due <b>' + dayMonth(i.due) + '</b> - ' + daysLate(i) + ' day' + (daysLate(i) === 1 ? '' : 's') + ' overdue' : 'due <b>' + dayMonth(i.due) + '</b>') + ')';
    const n = list.length;
    const count = (NUMBER_WORDS[n] || String(n));

    const contractText = contracts.length ? (contracts.length === 1 ? 'Contract ' + contracts[0] : 'Contracts ' + joinList(contracts, 'and')) : 'the contract';
    const subjectContracts = contracts.length ? (contracts.length === 1 ? 'Contract ' + contracts[0] : 'Contracts ' + joinList(contracts, '&')) : '';
    const subject = 'ACTION REQUIRED – Invoices awaiting authorisation' + (subjectContracts ? ' (' + subjectContracts + ')' : '');

    const P = 'margin:0 0 12px';
    const parts = [];
    parts.push('<p style="' + P + '">Hi all,</p>');
    parts.push('<p style="' + P + '">I\'m preparing ' + runDay + '\'s payment run, which will capture all invoices due up to <b>' + longDay(cutoff) + '</b>. ' +
      (n === 1 ? 'The invoice below on' : 'The ' + count + ' invoices below on') + ' <b>' + esc(contractText) + '</b> ' + (n === 1 ? 'is' : 'are') +
      ' awaiting authorisation before I can release ' + (n === 1 ? 'it' : 'them') + ' for payment.</p>');

    // the call-outs: red for invoices already overdue (top priority), orange for those that must make this run to avoid going overdue
    const box = (bg, border, html) => '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;margin:0 0 12px"><tr><td style="background:' + bg + ';border-left:4px solid ' + border + ';padding:10px 14px;font-size:14px;color:#1f2937">' + html + '</td></tr></table>';
    const callouts = [];
    if (overdue.length) {
      const sum = overdue.reduce((a, i) => a + i.balance, 0);
      callouts.push(box('#fdeaea', '#c62828',
        '<b>OVERDUE – top priority:</b> ' + (overdue.length === 1 ? 'one invoice is' : (NUMBER_WORDS[overdue.length] || overdue.length) + ' invoices (' + money(sum) + ') are') +
        ' already overdue: ' + overdue.map(i => describe(i, true)).join('; ') +
        '. If I can get authorisation by first thing ' + runDay + ' ' + (overdue.length === 1 ? 'it' : 'they') + ' can still go out in this run.'));
    }
    if (soon.length) {
      const sum = soon.reduce((a, i) => a + i.balance, 0);
      callouts.push(box('#fff4e5', '#e08a00',
        '<b>Priority:</b> ' + (soon.length === 1 ? 'one invoice is' : (NUMBER_WORDS[soon.length] || soon.length) + ' invoices (' + money(sum) + ') are') +
        ' not yet overdue but ' + (soon.length === 1 ? 'is' : 'are') + ' due on or before ' + dayMonth(cutoff) + ', so ' + (soon.length === 1 ? 'it needs' : 'they need') + ' to make ' + runDay + '\'s run: ' +
        soon.map(i => describe(i, false)).join('; ') +
        '. If I can get authorisation by first thing ' + runDay + ' ' + (soon.length === 1 ? 'it' : 'they') + ' can go out in this run; otherwise ' + (soon.length === 1 ? 'it' : 'they') + ' will go overdue.'));
    }
    if (!inside.length) {
      const first = list.find(i => i.due);
      callouts.push(box('#fff4e5', '#e08a00',
        '<b>Timing:</b> none of these are overdue and none fall due inside ' + runDay + '\'s window' +
        (first ? ', but the first is due on <b>' + dayMonth(first.due) + '</b>' : '') +
        '. Clearing them now means they\'re ready in good time for the next runs and nothing slips into overdue.'));
    }
    parts.push(callouts.join('\n'));

    // the table
    const cell = 'padding:8px 10px;border-top:1px solid #e3e8ef';
    const th = (t, right) => '<th style="padding:8px 10px;font-weight:600' + (right ? ';text-align:right' : '') + '">' + t + '</th>';
    let t = '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;font-size:13px;margin:0 0 16px;border:1px solid #cfd6e0">';
    t += '<tr style="background:#0f4c81;color:#ffffff;text-align:left">' + th('Due') + (multi ? th('Contract') : '') + th('Supplier') + th('PO') + th('Supplier inv. ref') + th('Invoice date') + th('Balance', true) + '</tr>';
    list.forEach((i, k) => {
      const bg = isOverdue(i) ? '#fdeaea' : isInside(i) ? '#fff4e5' : (k % 2 ? '#f5f8fc' : '#ffffff');
      t += '<tr style="background:' + bg + '">' +
        '<td style="' + cell + '"><b>' + (i.due ? shortDate(i.due) : '') + '</b></td>' +
        (multi ? '<td style="' + cell + '">' + esc(i.contract) + '</td>' : '') +
        '<td style="' + cell + '">' + esc(i.supplier) + '</td>' +
        '<td style="' + cell + '">' + esc(i.po) + '</td>' +
        '<td style="' + cell + '">' + esc(i.ref) + '</td>' +
        '<td style="' + cell + '">' + (i.invDate ? shortDate(i.invDate) : '') + '</td>' +
        '<td style="' + cell + ';text-align:right">' + money(i.balance) + '</td></tr>';
    });
    t += '<tr style="background:#e8eef6"><td colspan="' + (multi ? 6 : 5) + '" style="padding:9px 10px;border-top:2px solid #0f4c81;text-align:right"><b>Total awaiting authorisation</b></td>' +
      '<td style="padding:9px 10px;border-top:2px solid #0f4c81;text-align:right"><b>' + money(total) + '</b></td></tr></table>';
    parts.push(t);

    parts.push('<p style="margin:0 0 6px"><b>What I need from you</b></p>');
    parts.push('<p style="margin:0 0 14px">Please review these, authorise on COINS and return to finance. If there is a problem with any of them, let me know and I\'ll take it up with the supplier.</p>');
    if (inside.length === 0 && replyBy.getTime() > today.getTime()) {
      parts.push('<p style="' + P + '">If you could come back to me by <b>' + longDay(replyBy) + '</b>, that would keep everything on track for payment on time.</p>');
    }
    parts.push('<p style="' + P + '">Thanks very much for your help.</p>');
    parts.push('<p style="margin:0">Kind regards,<br><b>[Your name]</b><br><span style="color:#64748b">[Job title] | [Phone]</span></p>');

    const html = '<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#1f2937;line-height:1.5;max-width:700px">\n' + parts.join('\n') + '\n</div>';

    // plain-text version for anything that can't take formatting
    const text = ['Subject: ' + subject, ''].concat(
      list.map(i => [i.due ? shortDate(i.due) : '', multi ? i.contract : null, i.supplier, i.po, i.ref, i.invDate ? shortDate(i.invDate) : '', money(i.balance)].filter(x => x !== null).join('\t')),
      ['Total awaiting authorisation\t' + money(total)]).join('\n');

    return { subject, html, text, count: n, total, priority: inside.length };
  }

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = { toDate, parseRows, buildEmail, nextMonday, shortName, money };
  }
  if (typeof document === 'undefined') return;

  // ------------------------------------------------------------------
  // Page wiring
  // ------------------------------------------------------------------
  const zone = document.querySelector('.authDrop');
  if (!zone) return;
  const fileInput = zone.querySelector('input[type=file]');
  const fileLabel = zone.querySelector('.authFile');
  const hint = zone.querySelector('.authHint');
  const btn = document.querySelector('.authBtn');
  const subjBtn = document.querySelector('.authSubjectBtn');
  const dateInput = document.querySelector('.authRunDate');
  const status = document.querySelector('.authStatus');
  const preview = document.querySelector('.authPreview');
  const subjectBox = document.querySelector('.authSubject');
  const emailBox = document.querySelector('.authEmail');

  let file = null;
  let current = null;

  const iso = d => d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0') + '-' + String(d.getUTCDate()).padStart(2, '0');
  const now = new Date();
  const todayUtc = utc(now.getFullYear(), now.getMonth(), now.getDate());
  if (dateInput) dateInput.value = iso(nextMonday(now));

  function setStatus(msg, kind) {
    status.textContent = msg;
    status.classList.remove('ok', 'warn', 'bad');
    if (kind) status.classList.add(kind);
  }

  function setFile(f) {
    if (!f) return;
    if (!/\.xlsx?$/i.test(f.name)) { setStatus('That isn\'t a spreadsheet (.xlsx).', 'bad'); return; }
    file = f;
    current = null;
    fileLabel.textContent = f.name;
    hint.style.display = 'none';
    preview.style.display = 'none';
    setStatus('', '');
  }

  zone.addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('click', e => e.stopPropagation());
  fileInput.addEventListener('change', () => { setFile(fileInput.files[0]); fileInput.value = ''; });
  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('dragover'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.classList.remove('dragover');
    if (e.dataTransfer && e.dataTransfer.files.length) setFile(e.dataTransfer.files[0]);
  });

  // Copy text to the clipboard (used for the subject line)
  function copyPlain(text) {
    try {
      const ta = document.createElement('textarea');
      ta.value = text;
      ta.style.position = 'fixed'; ta.style.opacity = '0';
      document.body.appendChild(ta);
      ta.select();
      const ok = document.execCommand('copy');
      document.body.removeChild(ta);
      if (ok) return Promise.resolve(true);
    } catch (e) { /* fall through */ }
    if (navigator.clipboard && navigator.clipboard.writeText) return navigator.clipboard.writeText(text).then(() => true, () => false);
    return Promise.resolve(false);
  }

  // Copy the e-mail as formatted HTML, so Outlook pastes the table with its colours.
  // First choice: put the exact HTML on the clipboard.  Fallback: select the preview and use the browser's own copy.
  function copyRich(html, text) {
    if (navigator.clipboard && navigator.clipboard.write && typeof ClipboardItem !== 'undefined') {
      try {
        const item = new ClipboardItem({
          'text/html': new Blob([html], { type: 'text/html' }),
          'text/plain': new Blob([text], { type: 'text/plain' })
        });
        return navigator.clipboard.write([item]).then(() => true, () => selectAndCopy());
      } catch (e) { /* fall through */ }
    }
    return Promise.resolve(selectAndCopy());
  }

  function selectAndCopy() {
    try {
      const r = document.createRange();
      r.selectNodeContents(emailBox);
      const s = window.getSelection();
      s.removeAllRanges();
      s.addRange(r);
      const ok = document.execCommand('copy');
      s.removeAllRanges();
      return ok;
    } catch (e) { return false; }
  }

  async function generate() {
    if (!file) { setStatus('Drop the spreadsheet in the box first.', 'bad'); return; }
    if (typeof XLSX === 'undefined') { setStatus('The spreadsheet reader didn\'t load (needs an internet connection) - reload the page and try again.', 'bad'); return; }
    try {
      const wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: '' });
      const invoices = parseRows(rows);
      const run = (dateInput && toDate(dateInput.value)) || nextMonday(now);
      current = buildEmail(invoices, { runDate: run, today: todayUtc });

      subjectBox.textContent = current.subject;
      emailBox.innerHTML = current.html;
      preview.style.display = '';
      subjBtn.style.display = '';

      const ok = await copyRich(current.html, current.text);
      setStatus(ok
        ? 'Email copied (' + current.count + ' invoice' + (current.count > 1 ? 's' : '') + ', ' + money(current.total) + ') - paste it into a new Outlook message with Ctrl+V. Use "Copy subject" for the subject line.'
        : 'Email built below, but the browser wouldn\'t copy it - click inside it, press Ctrl+A then Ctrl+C.',
        ok ? 'ok' : 'warn');
    } catch (e) {
      console.error(e);
      setStatus(e && e.message ? e.message : 'Could not read that spreadsheet.', 'bad');
    }
  }

  btn.addEventListener('click', e => { e.preventDefault(); generate(); });
  subjBtn.addEventListener('click', async e => {
    e.preventDefault();
    if (!current) return;
    const ok = await copyPlain(current.subject);
    setStatus(ok ? 'Subject copied: ' + current.subject : 'Could not copy the subject - select it above and press Ctrl+C.', ok ? 'ok' : 'warn');
  });
})();
