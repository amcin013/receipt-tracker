/**
 * ============================================================================
 * Processor.gs — Inbox folder -> Gemini extraction -> Expenses row
 * ----------------------------------------------------------------------------
 * Created:  2026-10-30
 * Author:   Allen McIntosh (with Claude)
 *
 * Flow for each file dropped in "Expense Receipts/Inbox":
 *   1. Skip if already recorded (by Drive file ID)
 *   2. Validate type/size
 *   3. Extract fields with Gemini
 *   4. Rename file (date_vendor_total_ID) and move it to
 *      Processed/YYYY/MM  — or Needs Attention if it needs review
 *   5. Append a row to the Expenses tab with a link to the file
 *
 * Transient failures (rate limit, outage, bad key) leave the file in the
 * Inbox so the next run retries it. File-specific failures still get a row
 * (Status = Needs Review) so nothing silently disappears.
 * ============================================================================
 */

/**
 * Time-driven trigger entry point.
 * @param {Object=} e - trigger event (unused)
 * @return {{processed: number, review: number, skipped: number, stopped: string}}
 */
function processInbox(e) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10 * 1000)) {
    Logger.log('Another run is in progress; skipping.');
    return { processed: 0, review: 0, skipped: 0, stopped: 'Another run is already in progress.' };
  }

  const result = { processed: 0, review: 0, skipped: 0, stopped: '' };
  try {
    const folders = ensureFolders_();
    const ctx = buildContext_();
    const started = Date.now();
    const files = folders.inbox.getFiles();

    while (files.hasNext()) {
      if (Date.now() - started > CONFIG.MAX_RUNTIME_MS) {
        result.stopped = 'Time limit reached; remaining files will be picked up next run.';
        break;
      }
      const file = files.next();
      try {
        const outcome = processFile_(file, folders, ctx);
        if (outcome === 'ok') result.processed++;
        else if (outcome === 'review') result.review++;
        else result.skipped++;
      } catch (err) {
        if (err instanceof ExtractionError && err.retryable) {
          logEvent_('WARN', file.getName(), err.message + ' (will retry next run)');
          result.skipped++;
          if (err.fatal) { result.stopped = err.message; break; }
        } else {
          // Unexpected error: still record it so the receipt is not lost
          logEvent_('ERROR', file.getName(), String(err && err.stack || err));
          recordFailure_(file, folders, ctx, String(err && err.message || err));
          result.review++;
        }
      }
    }
  } finally {
    lock.releaseLock();
  }
  return result;
}

/**
 * Menu wrapper: runs processInbox and shows a summary.
 */
function processInboxFromMenu() {
  const r = processInbox();
  SpreadsheetApp.getUi().alert('Inbox processed',
    `Recorded: ${r.processed}\nNeeds review: ${r.review}\nSkipped / will retry: ${r.skipped}` +
    (r.stopped ? `\n\nNote: ${r.stopped}` : ''),
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/* ============================================================================
 * Per-run context (loaded once, reused for every file)
 * ========================================================================== */

/**
 * Loads categories and existing rows so each file does not re-read the sheet.
 * @return {Object}
 */
function buildContext_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const catSheet = ss.getSheetByName(CONFIG.SHEET_CATEGORIES);
  const expSheet = ss.getSheetByName(CONFIG.SHEET_EXPENSES);
  if (!catSheet || !expSheet) throw new Error('Run Expenses > Setup > Run first-time setup first.');

  const categories = catSheet.getLastRow() > 1
    ? catSheet.getRange(2, 1, catSheet.getLastRow() - 1, 4).getValues()
        .filter(r => String(r[0]).trim())
        .map(r => ({ name: String(r[0]).trim(), type: r[1] || TYPES.TRACK_ONLY, description: String(r[3] || '') }))
    : [];
  if (!categories.some(c => c.name === 'Other')) {
    categories.push({ name: 'Other', type: TYPES.TRACK_ONLY, description: 'Anything else' });
  }

  // Index existing rows for duplicate detection
  const fileIds = new Set();
  const dupKeys = {};
  if (expSheet.getLastRow() > 1) {
    const data = expSheet.getRange(2, 1, expSheet.getLastRow() - 1, COLUMNS.length).getValues();
    data.forEach(r => {
      const fid = r[col_('fileId') - 1];
      if (fid) fileIds.add(String(fid));
      const key = dupKey_(r[col_('vendor') - 1], r[col_('date') - 1], r[col_('total') - 1]);
      if (key) dupKeys[key] = r[col_('id') - 1];
    });
  }

  return { sheet: expSheet, categories: categories, fileIds: fileIds, dupKeys: dupKeys };
}

/* ============================================================================
 * Single-file processing
 * ========================================================================== */

/**
 * Processes one Inbox file.
 * @param {File} file
 * @param {Object} folders - from ensureFolders_()
 * @param {Object} ctx - from buildContext_()
 * @return {string} 'ok' | 'review' | 'skip'
 */
function processFile_(file, folders, ctx) {
  const name = file.getName();

  // 1. Already recorded? (e.g., a previous run wrote the row but failed to move the file)
  if (ctx.fileIds.has(file.getId())) {
    file.moveTo(folders.processed);
    logEvent_('INFO', name, 'Already recorded; moved out of Inbox.');
    return 'skip';
  }

  // 2. Validate
  const mime = file.getMimeType();
  if (CONFIG.SUPPORTED_MIME_TYPES.indexOf(mime) < 0) {
    recordFailure_(file, folders, ctx,
      'Unsupported file type (' + mime + '). Use PDF, JPG, PNG, WEBP, or HEIC.');
    return 'review';
  }
  if (file.getSize() > CONFIG.MAX_FILE_BYTES) {
    recordFailure_(file, folders, ctx, 'File is larger than 15 MB. Re-scan at a lower quality.');
    return 'review';
  }

  // 3. Extract
  let data;
  try {
    data = extractReceiptData_(file.getBlob(), ctx.categories);
  } catch (err) {
    if (err instanceof ExtractionError && !err.retryable) {
      recordFailure_(file, folders, ctx, err.message);
      return 'review';
    }
    throw err; // retryable -> handled by caller
  }

  // 4. Normalize fields
  const fields = normalizeFields_(data, ctx.categories);
  const id = nextReceiptId_();
  const notes = [];
  if (!data.is_receipt) notes.push('Does not look like a receipt.');
  if (data.confidence === 'low') notes.push('Low confidence read; please verify.');
  if (fields.total === '') notes.push('Total not found.');
  if (fields.date === '') notes.push('Date not found.');
  if (data.warnings) notes.push(String(data.warnings));
  if (data.currency && String(data.currency).toUpperCase() !== 'USD') notes.push('Currency: ' + data.currency);

  const key = dupKey_(fields.vendor, fields.date, fields.total);
  if (key && ctx.dupKeys[key]) notes.push('Possible duplicate of ' + ctx.dupKeys[key] + '.');

  const needsReview = !data.is_receipt || data.confidence === 'low' ||
                      fields.total === '' || fields.date === '' || (key && ctx.dupKeys[key]);

  // 5. Rename and file it
  const newName = buildFileName_(fields, id, name);
  file.setName(newName);
  const target = needsReview ? folders.attention : monthFolder_(folders.processed, fields.date || new Date());
  file.moveTo(target);

  // 6. Status and purpose defaults
  let status;
  if (needsReview) status = STATUS.NEEDS_REVIEW;
  else if (fields.type === TYPES.REIMBURSABLE) status = STATUS.UNSUBMITTED;
  else status = STATUS.TRACKED;

  let purpose = '';
  if (fields.category === 'Fuel') purpose = getSetting_('Default Fuel Purpose');

  // 7. Write the row
  writeRow_(ctx.sheet, {
    select: false,
    id: id,
    added: new Date(),
    date: fields.date,
    vendor: fields.vendor,
    category: fields.category,
    type: fields.type,
    subtotal: fields.subtotal,
    tax: fields.tax,
    total: fields.total,
    payment: fields.payment,
    gallons: fields.gallons,
    ppg: fields.ppg,
    purpose: purpose,
    site: '',
    status: status,
    submitted: '',
    batch: '',
    receipt: '=HYPERLINK("' + file.getUrl() + '","View")',
    items: fields.items,
    notes: notes.join(' '),
    fileId: file.getId()
  });

  ctx.fileIds.add(file.getId());
  if (key) ctx.dupKeys[key] = id;
  logEvent_('INFO', newName, `Recorded ${id}: ${fields.vendor} ${money_(fields.total)} [${fields.category}]` +
    (needsReview ? ' (needs review)' : ''));
  return needsReview ? 'review' : 'ok';
}

/**
 * Records a file that could not be read so it still shows up on the sheet.
 * @param {File} file
 * @param {Object} folders
 * @param {Object} ctx
 * @param {string} reason
 */
function recordFailure_(file, folders, ctx, reason) {
  const id = nextReceiptId_();
  const original = file.getName();
  try {
    file.setName(id + '_' + original);
    file.moveTo(folders.attention);
  } catch (e) {
    reason += ' (Also failed to move file: ' + e.message + ')';
  }
  writeRow_(ctx.sheet, {
    select: false, id: id, added: new Date(), date: '', vendor: '', category: '',
    type: '', subtotal: '', tax: '', total: '', payment: '', gallons: '', ppg: '',
    purpose: '', site: '', status: STATUS.NEEDS_REVIEW, submitted: '', batch: '',
    receipt: '=HYPERLINK("' + file.getUrl() + '","View")',
    items: '', notes: 'Could not read automatically: ' + reason, fileId: file.getId()
  });
  ctx.fileIds.add(file.getId());
  logEvent_('WARN', original, `Recorded ${id} as Needs Review: ${reason}`);
}

/* ============================================================================
 * Helpers
 * ========================================================================== */

/**
 * Cleans Gemini output into sheet-ready values.
 * @param {Object} d - raw extraction
 * @param {{name: string, type: string}[]} categories
 * @return {Object}
 */
function normalizeFields_(d, categories) {
  const num = v => (v === null || v === undefined || v === '' || isNaN(Number(v))) ? '' : Math.round(Number(v) * 1000) / 1000;
  const cat = categories.find(c => c.name === d.category) || categories.find(c => c.name === 'Other');

  // Card last 4: digits only, max 4 (guards against a full number slipping through)
  let last4 = String(d.payment_last4 || '').replace(/\D/g, '');
  last4 = last4 ? last4.slice(-4) : '';

  let items = d.items_summary || '';
  if (d.fuel_grade) items = (items ? items + ' | ' : '') + 'Grade: ' + d.fuel_grade;

  return {
    date: parseIsoDate_(d.receipt_date),
    vendor: safeText_(String(d.vendor || '').trim()),
    category: cat.name,
    type: cat.type,
    subtotal: num(d.subtotal),
    tax: num(d.tax),
    total: num(d.total) === '' ? '' : Math.round(Number(d.total) * 100) / 100,
    payment: last4,
    gallons: num(d.gallons),
    ppg: num(d.price_per_gallon),
    items: safeText_(items)
  };
}

/**
 * Parses YYYY-MM-DD into a local Date (no timezone drift). Rejects dates
 * more than a week in the future or before 2000 as likely misreads.
 * @param {string} s
 * @return {Date|string} Date, or '' if invalid
 */
function parseIsoDate_(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(s || ''));
  if (!m) return '';
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  if (isNaN(d.getTime()) || d.getMonth() !== Number(m[2]) - 1) return '';
  const weekAhead = new Date(Date.now() + 7 * 24 * 3600 * 1000);
  if (d.getFullYear() < 2000 || d > weekAhead) return '';
  return d;
}

/**
 * Generates the next receipt ID, e.g. R-2026-0042. Counter is per year.
 * Called only while the script lock is held.
 * @return {string}
 */
function nextReceiptId_() {
  const props = PropertiesService.getScriptProperties();
  const year = new Date().getFullYear();
  const raw = props.getProperty(CONFIG.PROP_COUNTER) || '';
  const [savedYear, savedN] = raw.split(':');
  const n = (Number(savedYear) === year ? Number(savedN) || 0 : 0) + 1;
  props.setProperty(CONFIG.PROP_COUNTER, year + ':' + n);
  return 'R-' + year + '-' + String(n).padStart(4, '0');
}

/**
 * Builds a descriptive file name: 2026-10-30_Shell_42.17_R-2026-0001.pdf
 * @param {Object} f - normalized fields
 * @param {string} id
 * @param {string} originalName
 * @return {string}
 */
function buildFileName_(f, id, originalName) {
  const ext = (/\.[a-z0-9]{2,5}$/i.exec(originalName) || [''])[0] || '';
  const datePart = f.date
    ? Utilities.formatDate(f.date, Session.getScriptTimeZone(), 'yyyy-MM-dd')
    : 'nodate';
  const vendor = String(f.vendor || 'Unknown').replace(/^'/, '')
    .replace(/[^A-Za-z0-9]+/g, '-').replace(/^-|-$/g, '').substring(0, 30) || 'Unknown';
  const total = f.total === '' ? 'nototal' : Number(f.total).toFixed(2);
  return `${datePart}_${vendor}_${total}_${id}${ext}`;
}

/**
 * Returns Processed/YYYY/MM for a date, creating folders as needed.
 * @param {Folder} processed
 * @param {Date} date
 * @return {Folder}
 */
function monthFolder_(processed, date) {
  const tz = Session.getScriptTimeZone();
  const y = getOrCreateSubfolder_(processed, Utilities.formatDate(date, tz, 'yyyy'));
  return getOrCreateSubfolder_(y, Utilities.formatDate(date, tz, 'MM'));
}

/**
 * Duplicate key: vendor + date + total. Returns '' if any part is missing.
 * @param {*} vendor
 * @param {*} date
 * @param {*} total
 * @return {string}
 */
function dupKey_(vendor, date, total) {
  if (!vendor || !(date instanceof Date) || total === '' || total === null) return '';
  const v = String(vendor).toLowerCase().replace(/[^a-z0-9]/g, '');
  const d = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyyMMdd');
  return v + '|' + d + '|' + Number(total).toFixed(2);
}

/**
 * Appends a row built from a {key: value} object and adds its checkbox.
 * @param {Sheet} sheet
 * @param {Object} obj
 */
function writeRow_(sheet, obj) {
  const row = COLUMNS.map(c => (obj[c.key] === undefined ? '' : obj[c.key]));
  sheet.appendRow(row);
  const r = sheet.getLastRow();
  sheet.getRange(r, col_('select')).insertCheckboxes();
}
