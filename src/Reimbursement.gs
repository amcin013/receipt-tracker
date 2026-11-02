/**
 * ============================================================================
 * Reimbursement.gs — Build reimbursement emails from selected rows
 * ----------------------------------------------------------------------------
 * Created:  2026-10-30
 * Updated:  2026-11-02  v2.0 — Work book only; optional Send From Alias
 * Author:   Allen McIntosh (with Claude)
 *
 * Work book only. Tick the "Select" checkbox on Work Expenses rows, then run
 * Expenses > Work: Prepare reimbursement email. One email is built with:
 *   - a summary table (date, vendor, purpose, gallons, amount) + grand total
 *   - every selected receipt attached, using its descriptive file name
 * Send Mode on the Settings tab decides whether it becomes a Gmail draft
 * (default: you review and hit Send) or is sent right away.
 * Rows are then stamped Submitted with a Batch ID so nothing is sent twice.
 * ============================================================================
 */

/**
 * Menu entry: validates selected rows and creates/sends the email.
 */
function prepareReimbursement() {
  const ui = SpreadsheetApp.getUi();
  const book = getBook_('work');
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(book.expensesSheet);
  const selected = getSelectedRows_(book, sh);

  if (!selected.length) {
    ui.alert('No rows selected', 'Tick the "Select" checkbox on the ' + book.expensesSheet + ' rows you want to submit.', ui.ButtonSet.OK);
    return;
  }

  // ---- Settings ------------------------------------------------------------
  const to = getSetting_('Reimbursement Email To');
  const cc = getSetting_('Reimbursement Email CC');
  const yourName = getSetting_('Your Name');
  const prefix = getSetting_('Subject Prefix') || 'Expense Reimbursement Request';
  const sendNow = getSetting_('Send Mode').toLowerCase() === 'send';
  const requirePurpose = settingIsYes_('Require Business Purpose');
  const fromAlias = getSetting_('Send From Alias');

  if (!to) {
    ui.alert('Missing recipient', 'Set "Reimbursement Email To" on the Settings tab first.', ui.ButtonSet.OK);
    return;
  }

  // Gmail only lets a script send "from" an address already verified under
  // Gmail > Settings > Accounts > "Send mail as". Check before building anything.
  if (fromAlias) {
    // getAliases() leaves out the account's own address, so add it explicitly
    const aliases = GmailApp.getAliases().concat([Session.getEffectiveUser().getEmail()])
      .map(a => String(a).toLowerCase());
    if (aliases.indexOf(fromAlias.toLowerCase()) < 0) {
      ui.alert('Send From Alias not set up',
        `"${fromAlias}" is not a verified "Send mail as" address in this Gmail account.\n\n` +
        'Add it in Gmail > Settings > Accounts > "Send mail as", or clear the Send From Alias setting.',
        ui.ButtonSet.OK);
      return;
    }
  }

  // ---- Validate rows -------------------------------------------------------
  const problems = [];
  selected.forEach(r => {
    const v = r.values;
    const label = v.id || ('row ' + r.row);
    if (v.type !== TYPES.REIMBURSABLE) problems.push(`${label}: Type is "${v.type || 'blank'}", not Reimbursable.`);
    if (v.status === STATUS.SUBMITTED || v.status === STATUS.REIMBURSED) {
      problems.push(`${label}: already ${v.status} (batch ${v.batch || 'n/a'}).`);
    }
    if (v.status === STATUS.NEEDS_REVIEW) problems.push(`${label}: still marked Needs Review. Verify it and change Status to Unsubmitted.`);
    if (typeof v.total !== 'number' || v.total <= 0) problems.push(`${label}: Total is missing or invalid.`);
    if (!(v.date instanceof Date)) problems.push(`${label}: Receipt Date is missing.`);
    if (requirePurpose && !String(v.purpose).trim()) problems.push(`${label}: Business Purpose is blank.`);
    if (!v.fileId) problems.push(`${label}: no receipt file linked.`);
  });
  if (problems.length) {
    ui.alert('Fix these before submitting', problems.join('\n'), ui.ButtonSet.OK);
    return;
  }

  // ---- Gather attachments --------------------------------------------------
  const attachments = [];
  let bytes = 0;
  for (const r of selected) {
    try {
      const file = DriveApp.getFileById(r.values.fileId);
      const blob = file.getBlob().setName(file.getName());
      bytes += blob.getBytes().length;
      attachments.push(blob);
    } catch (e) {
      ui.alert('Receipt file missing', `Could not open the file for ${r.values.id}: ${e.message}`, ui.ButtonSet.OK);
      return;
    }
  }
  if (bytes > CONFIG.MAX_EMAIL_ATTACH_BYTES) {
    ui.alert('Attachments too large',
      `Selected receipts total ${(bytes / 1048576).toFixed(1)} MB (Gmail limit is 25 MB). Select fewer rows and send in two batches.`,
      ui.ButtonSet.OK);
    return;
  }

  // ---- Build email ---------------------------------------------------------
  const tz = Session.getScriptTimeZone();
  const fmt = d => Utilities.formatDate(d, tz, 'MM/dd/yyyy');
  const sorted = selected.slice().sort((a, b) => a.values.date - b.values.date);
  const grand = sorted.reduce((s, r) => s + r.values.total, 0);
  const first = fmt(sorted[0].values.date);
  const last = fmt(sorted[sorted.length - 1].values.date);
  const range = first === last ? first : `${first} - ${last}`;
  const subject = `${prefix}: ${sorted.length} receipt${sorted.length > 1 ? 's' : ''}, ${money_(grand)} (${range})`;
  const batchId = 'B-' + Utilities.formatDate(new Date(), tz, 'yyyyMMdd-HHmmss');

  const html = buildEmailHtml_(sorted, grand, range, batchId, yourName, fmt);
  const plain = buildEmailPlain_(sorted, grand, range, batchId, yourName, fmt);

  const confirm = ui.alert('Confirm',
    `${sendNow ? 'SEND' : 'Create a Gmail DRAFT of'} this email?\n\n` +
    (fromAlias ? `From: ${fromAlias}\n` : '') +
    `To: ${to}${cc ? '\nCC: ' + cc : ''}\nSubject: ${subject}\nAttachments: ${attachments.length}`,
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  const options = { htmlBody: html, attachments: attachments };
  if (cc) options.cc = cc;
  if (yourName) options.name = yourName;
  if (fromAlias) options.from = fromAlias;

  if (sendNow) GmailApp.sendEmail(to, subject, plain, options);
  else GmailApp.createDraft(to, subject, plain, options);

  // ---- Stamp rows ----------------------------------------------------------
  const today = new Date(); today.setHours(0, 0, 0, 0);
  selected.forEach(r => {
    sh.getRange(r.row, col_(book, 'status')).setValue(STATUS.SUBMITTED);
    sh.getRange(r.row, col_(book, 'submitted')).setValue(today);
    sh.getRange(r.row, col_(book, 'batch')).setValue(batchId);
    sh.getRange(r.row, col_(book, 'select')).setValue(false);
  });

  logEvent_('INFO', batchId, `${sendNow ? 'Sent' : 'Drafted'} reimbursement: ${selected.length} receipts, ${money_(grand)} to ${to}`);
  ui.alert(sendNow ? 'Sent' : 'Draft created',
    sendNow
      ? `Reimbursement email sent. Batch ${batchId}.`
      : `A draft is waiting in Gmail > Drafts. Review it and press Send.\n\n` +
        `Rows are marked Submitted (batch ${batchId}). If you discard the draft, use ` +
        `Expenses > Work: Undo a reimbursement batch to reset them.`,
    ui.ButtonSet.OK);
}

/**
 * Menu entry: marks selected Submitted rows as Reimbursed.
 */
function markSelectedReimbursed() {
  const ui = SpreadsheetApp.getUi();
  const book = getBook_('work');
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(book.expensesSheet);
  const selected = getSelectedRows_(book, sh);
  if (!selected.length) { ui.alert('Tick the "Select" checkbox on the paid rows first.'); return; }

  let changed = 0;
  const skipped = [];
  selected.forEach(r => {
    if (r.values.status === STATUS.SUBMITTED) {
      sh.getRange(r.row, col_(book, 'status')).setValue(STATUS.REIMBURSED);
      changed++;
    } else {
      skipped.push(`${r.values.id || 'row ' + r.row} (${r.values.status || 'no status'})`);
    }
    sh.getRange(r.row, col_(book, 'select')).setValue(false);
  });
  ui.alert(`Marked ${changed} row(s) Reimbursed.` +
    (skipped.length ? `\n\nSkipped (not Submitted):\n${skipped.join('\n')}` : ''));
}

/**
 * Menu entry: resets every row in a batch back to Unsubmitted.
 * Use when a draft was discarded or a submission needs to be redone.
 * Does not delete any email or draft.
 */
function undoReimbursementBatch() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('Undo batch', 'Enter the Batch ID to reset (e.g., B-20261030-101500):', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const batchId = res.getResponseText().trim();
  if (!batchId) return;

  const book = getBook_('work');
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(book.expensesSheet);
  if (sh.getLastRow() < 2) return;
  const batches = sh.getRange(2, col_(book, 'batch'), sh.getLastRow() - 1, 1).getValues();

  let count = 0;
  batches.forEach((b, i) => {
    if (String(b[0]).trim() === batchId) {
      const row = i + 2;
      sh.getRange(row, col_(book, 'status')).setValue(STATUS.UNSUBMITTED);
      sh.getRange(row, col_(book, 'submitted')).clearContent();
      sh.getRange(row, col_(book, 'batch')).clearContent();
      count++;
    }
  });
  logEvent_('INFO', batchId, `Batch undone; ${count} row(s) reset to Unsubmitted.`);
  ui.alert(count ? `Reset ${count} row(s) to Unsubmitted.` : 'No rows found with that Batch ID.');
}

/* ============================================================================
 * Helpers
 * ========================================================================== */

/**
 * Returns rows whose Select checkbox is ticked, as {row, values{key: value}}.
 * @param {Object} book
 * @param {Sheet} sh
 * @return {{row: number, values: Object}[]}
 */
function getSelectedRows_(book, sh) {
  if (sh.getLastRow() < 2) return [];
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, book.columns.length).getValues();
  const sel = col_(book, 'select') - 1;
  const out = [];
  data.forEach((r, i) => {
    if (r[sel] === true) out.push({ row: i + 2, values: rowToObject_(book, r) });
  });
  return out;
}

/**
 * HTML body: intro, table, total, sign-off.
 */
function buildEmailHtml_(rows, grand, range, batchId, yourName, fmt) {
  const th = 'style="text-align:left;padding:6px 10px;background:#1f3864;color:#fff;font-weight:bold"';
  const td = 'style="padding:6px 10px;border-bottom:1px solid #ddd"';
  const tdr = 'style="padding:6px 10px;border-bottom:1px solid #ddd;text-align:right"';
  const hasFuel = rows.some(r => r.values.gallons !== '' && r.values.gallons !== null);

  const body = rows.map(r => {
    const v = r.values;
    return '<tr>' +
      `<td ${td}>${escapeHtml_(fmt(v.date))}</td>` +
      `<td ${td}>${escapeHtml_(v.vendor)}</td>` +
      `<td ${td}>${escapeHtml_(v.category)}</td>` +
      `<td ${td}>${escapeHtml_(v.purpose)}${v.site ? ' (' + escapeHtml_(v.site) + ')' : ''}</td>` +
      (hasFuel ? `<td ${tdr}>${v.gallons === '' ? '' : escapeHtml_(Number(v.gallons).toFixed(3))}</td>` : '') +
      `<td ${tdr}>${escapeHtml_(money_(v.total))}</td>` +
      `<td ${td}>${escapeHtml_(v.id)}</td>` +
      '</tr>';
  }).join('');

  return '<div style="font-family:Arial,sans-serif;font-size:14px;color:#222">' +
    '<p>Hello,</p>' +
    `<p>Please find attached ${rows.length} receipt${rows.length > 1 ? 's' : ''} for reimbursement ` +
    `covering ${escapeHtml_(range)}, totaling <b>${escapeHtml_(money_(grand))}</b>.</p>` +
    '<table style="border-collapse:collapse;font-size:13px">' +
    `<tr><th ${th}>Date</th><th ${th}>Vendor</th><th ${th}>Category</th><th ${th}>Business Purpose</th>` +
    (hasFuel ? `<th ${th}>Gallons</th>` : '') +
    `<th ${th}>Amount</th><th ${th}>Receipt</th></tr>` +
    body +
    `<tr><td colspan="${hasFuel ? 5 : 4}" style="padding:8px 10px;text-align:right;font-weight:bold">Total</td>` +
    `<td style="padding:8px 10px;text-align:right;font-weight:bold">${escapeHtml_(money_(grand))}</td><td></td></tr>` +
    '</table>' +
    '<p>Each attachment is named with its date, vendor, amount, and receipt number to match the table above.</p>' +
    `<p>Thank you,<br>${escapeHtml_(yourName)}</p>` +
    `<p style="color:#888;font-size:11px">Reference: ${escapeHtml_(batchId)}</p>` +
    '</div>';
}

/**
 * Plain-text fallback body.
 */
function buildEmailPlain_(rows, grand, range, batchId, yourName, fmt) {
  const lines = rows.map(r => {
    const v = r.values;
    return `${fmt(v.date)}  ${v.vendor}  ${v.category}  ${money_(v.total)}  ${v.purpose}  [${v.id}]`;
  });
  return 'Hello,\n\n' +
    `Please find attached ${rows.length} receipt(s) for reimbursement covering ${range}, totaling ${money_(grand)}.\n\n` +
    lines.join('\n') + `\n\nTotal: ${money_(grand)}\n\nThank you,\n${yourName}\n\nReference: ${batchId}`;
}
