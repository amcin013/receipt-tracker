/**
 * ============================================================================
 * Processor.gs — Inbox folders -> Gemini extraction -> book rows
 * ----------------------------------------------------------------------------
 * Created:  2026-10-30
 * Updated:  2026-11-02  v2.0 — processes Work and Home inboxes, archives
 *                       untouched originals, writes the row before moving
 * Author:   Allen McIntosh (with Claude)
 *
 * Flow for each file dropped in "Receipts/<Book>/Inbox":
 *   1. Skip if already recorded (by Drive file ID)
 *   2. Validate type/size
 *   3. Extract fields with Gemini (using that book's categories)
 *   4. Copy the untouched file to <Book>/Originals/YYYY/MM
 *   5. Append a row to the book's Expenses tab with links to both files
 *   6. Rename the working file (date_vendor_total_ID) and move it to
 *      <Book>/Processed/YYYY/MM, or Needs Attention if it needs review
 *
 * The row is written BEFORE the file is moved. If the move fails, the row
 * already exists and the next run finishes the move (step 1), so a receipt
 * can never end up filed with no row.
 *
 * Transient failures (rate limit, outage, bad key) leave the file in the
 * Inbox so the next run retries it. File-specific failures still get a row
 * (Status = Needs Review) so nothing silently disappears.
 * ============================================================================
 */

/**
 * Time-driven trigger entry point. Processes the Work Inbox, then Home.
 * @param {Object=} e - trigger event (unused)
 * @return {{processed: number, review: number, skipped: number, stopped: string,
 *           byBook: Object<string, number>}}
 */
function processInbox(e) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10 * 1000)) {
    Logger.log('Another run is in progress; skipping.');
    return { processed: 0, review: 0, skipped: 0, stopped: 'Another run is already in progress.', byBook: {} };
  }

  const result = { processed: 0, review: 0, skipped: 0, stopped: '', byBook: {} };
  try {
    const folders = ensureFolders_();
    const ctx = buildContext_();
    const started = Date.now();

    for (const book of getBooks_()) {
      if (result.stopped) break;
      result.byBook[book.label] = 0;
      const bf = folders.books[book.key];
      const files = bf.inbox.getFiles();

      while (files.hasNext()) {
        if (Date.now() - started > CONFIG.MAX_RUNTIME_MS) {
          result.stopped = 'Time limit reached; remaining files will be picked up next run.';
          break;
        }
        const file = files.next();
        const where = book.label + ': ' + file.getName();
        try {
          const outcome = processFile_(book, file, bf, ctx);
          if (outcome === 'ok') result.processed++;
          else if (outcome === 'review') result.review++;
          else result.skipped++;
          if (outcome !== 'skip') result.byBook[book.label]++;
        } catch (err) {
          if (err instanceof ExtractionError && err.retryable) {
            logEvent_('WARN', where, err.message + ' (will retry next run)');
            result.skipped++;
            if (err.fatal) { result.stopped = err.message; break; }
          } else {
            // Unexpected error: still record it so the receipt is not lost
            logEvent_('ERROR', where, String(err && err.stack || err));
            try {
              recordFailure_(book, file, bf, ctx, String(err && err.message || err));
              result.review++;
              result.byBook[book.label]++;
            } catch (err2) {
              // Could not even write a failure row (e.g., sheet missing).
              // Leave the file where it is; it will be retried next run.
              logEvent_('ERROR', where, 'Could not record failure: ' + err2);
              result.skipped++;
            }
          }
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
  const perBook = Object.keys(r.byBook).map(k => `${k}: ${r.byBook[k]}`).join('   ');
  SpreadsheetApp.getUi().alert('Inboxes processed',
    `Recorded: ${r.processed}\nNeeds review: ${r.review}\nSkipped / will retry: ${r.skipped}` +
    (perBook ? `\n\nNew rows by book -> ${perBook}` : '') +
    (r.stopped ? `\n\nNote: ${r.stopped}` : ''),
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/* ============================================================================
 * Per-run context (loaded once, reused for every file)
 * ========================================================================== */

/**
 * Loads each book's sheet and categories, plus an index of every recorded
 * file and receipt across BOTH books, so a receipt scanned into the wrong
 * inbox (or into both) is flagged as a possible duplicate.
 * @return {{books: Object<string, {sheet: Sheet, categories: Object[]}>,
 *           fileIds: Set<string>, dupKeys: Object<string, string>}}
 */
function buildContext_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ctx = { books: {}, fileIds: new Set(), dupKeys: {} };

  getBooks_().forEach(book => {
    const catSheet = ss.getSheetByName(book.categoriesSheet);
    const expSheet = ss.getSheetByName(book.expensesSheet);
    if (!catSheet || !expSheet) throw new Error('Run Expenses > Setup > Run first-time setup first.');

    ctx.books[book.key] = { sheet: expSheet, categories: loadCategories_(book, catSheet) };

    if (expSheet.getLastRow() > 1) {
      const data = expSheet.getRange(2, 1, expSheet.getLastRow() - 1, book.columns.length).getValues();
      data.forEach(r => {
        const v = rowToObject_(book, r);
        if (v.fileId) ctx.fileIds.add(String(v.fileId));
        const key = dupKey_(v.vendor, v.date, v.total);
        if (key) ctx.dupKeys[key] = v.id;
      });
    }
  });
  return ctx;
}

/**
 * Reads a book's Categories tab into {name, type, description} objects.
 * Always includes an "Other" fallback.
 * @param {Object} book
 * @param {Sheet} sh
 * @return {{name: string, type: string, description: string}[]}
 */
function loadCategories_(book, sh) {
  const cols = book.categoryColumns;
  const idx = key => cols.findIndex(c => c.key === key);
  const iType = idx('type'), iDesc = idx('description');

  const cats = sh.getLastRow() > 1
    ? sh.getRange(2, 1, sh.getLastRow() - 1, cols.length).getValues()
        .filter(r => String(r[0]).trim())
        .map(r => ({
          name: String(r[0]).trim(),
          type: iType >= 0 ? (r[iType] || TYPES.TRACK_ONLY) : '',
          description: iDesc >= 0 ? String(r[iDesc] || '') : ''
        }))
    : [];
  if (!cats.some(c => c.name === 'Other')) {
    cats.push({ name: 'Other', type: iType >= 0 ? TYPES.TRACK_ONLY : '', description: 'Anything else' });
  }
  return cats;
}

/* ============================================================================
 * Single-file processing
 * ========================================================================== */

/**
 * Processes one Inbox file for a book.
 * @param {Object} book
 * @param {File} file
 * @param {Object} bf  - this book's folders, from ensureFolders_().books[key]
 * @param {Object} ctx - from buildContext_()
 * @return {string} 'ok' | 'review' | 'skip'
 */
function processFile_(book, file, bf, ctx) {
  const originalName = file.getName();
  const where = book.label + ': ' + originalName;
  const bctx = ctx.books[book.key];

  // 1. Already recorded? (a previous run wrote the row but the move failed)
  if (ctx.fileIds.has(file.getId())) {
    file.moveTo(bf.processed);
    logEvent_('INFO', where, 'Already recorded; moved out of Inbox.');
    return 'skip';
  }

  // 2. Validate
  const mime = file.getMimeType();
  if (CONFIG.SUPPORTED_MIME_TYPES.indexOf(mime) < 0) {
    recordFailure_(book, file, bf, ctx,
      'Unsupported file type (' + mime + '). Use PDF, JPG, PNG, WEBP, or HEIC.');
    return 'review';
  }
  if (file.getSize() > CONFIG.MAX_FILE_BYTES) {
    recordFailure_(book, file, bf, ctx, 'File is larger than 15 MB. Re-scan at a lower quality.');
    return 'review';
  }

  // 3. Extract
  let data;
  try {
    data = extractReceiptData_(file.getBlob(), bctx.categories, book.aiContext);
  } catch (err) {
    if (err instanceof ExtractionError && !err.retryable) {
      recordFailure_(book, file, bf, ctx, err.message);
      return 'review';
    }
    throw err; // retryable -> handled by caller
  }

  // 4. Normalize fields and decide whether a human should look at it
  const fields = normalizeFields_(data, bctx.categories, book);
  const id = nextReceiptId_(book);
  const notes = [];
  if (!data.is_receipt) notes.push('Does not look like a receipt.');
  if (data.confidence === 'low') notes.push('Low confidence read; please verify.');
  if (fields.total === '') notes.push('Total not found.');
  if (fields.date === '') notes.push('Date not found.');
  if (data.warnings) notes.push(String(data.warnings));
  if (data.currency && String(data.currency).toUpperCase() !== 'USD') notes.push('Currency: ' + data.currency);

  const key = dupKey_(fields.vendor, fields.date, fields.total);
  const dupOf = key ? ctx.dupKeys[key] : '';
  if (dupOf) notes.push('Possible duplicate of ' + dupOf + '.');

  const needsReview = !data.is_receipt || data.confidence === 'low' ||
                      fields.total === '' || fields.date === '' || !!dupOf;

  // 5. Keep an untouched copy before anything is renamed
  const fileDate = fields.date || new Date();
  const original = archiveOriginal_(file, bf, id, originalName, fileDate, notes);

  // 6. Status and purpose defaults
  const status = needsReview ? STATUS.NEEDS_REVIEW : book.initialStatus(fields.type);
  let purpose = '';
  if (book.key === 'work' && fields.category === 'Fuel') purpose = getSetting_('Default Fuel Purpose');

  // 7. Write the row (Drive URLs do not change when a file is moved/renamed)
  writeRow_(book, bctx.sheet, {
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
    memo: '',
    status: status,
    submitted: '',
    batch: '',
    receipt: linkFormula_(file.getUrl()),
    original: original ? linkFormula_(original.getUrl()) : '',
    items: fields.items,
    notes: safeText_(notes.join(' ')),
    fileId: file.getId(),
    originalId: original ? original.getId() : ''
  });
  ctx.fileIds.add(file.getId());
  if (key) ctx.dupKeys[key] = id;

  // 8. Rename and file the working copy. A failure here is not fatal: the
  //    row exists, and step 1 moves the file on the next run.
  const newName = buildFileName_(fields, id, originalName);
  try {
    file.setName(newName);
    file.moveTo(needsReview ? bf.attention : monthFolder_(bf.processed, fileDate));
  } catch (e) {
    logEvent_('WARN', where, `Recorded ${id} but could not rename/move the file: ${e.message}. Will retry next run.`);
  }

  logEvent_('INFO', book.label + ': ' + newName,
    `Recorded ${id}: ${fields.vendor} ${money_(fields.total)} [${fields.category}]` +
    (needsReview ? ' (needs review)' : ''));
  return needsReview ? 'review' : 'ok';
}

/**
 * Records a file that could not be read so it still shows up on the sheet.
 * @param {Object} book
 * @param {File} file
 * @param {Object} bf - this book's folders
 * @param {Object} ctx
 * @param {string} reason
 */
function recordFailure_(book, file, bf, ctx, reason) {
  const id = nextReceiptId_(book);
  const originalName = file.getName();
  const extra = [];
  const original = archiveOriginal_(file, bf, id, originalName, new Date(), extra);

  writeRow_(book, ctx.books[book.key].sheet, {
    select: false, id: id, added: new Date(), status: STATUS.NEEDS_REVIEW,
    receipt: linkFormula_(file.getUrl()),
    original: original ? linkFormula_(original.getUrl()) : '',
    notes: safeText_('Could not read automatically: ' + reason + (extra.length ? ' ' + extra.join(' ') : '')),
    fileId: file.getId(),
    originalId: original ? original.getId() : ''
  });
  ctx.fileIds.add(file.getId());

  try {
    file.setName(id + '_' + originalName);
    file.moveTo(bf.attention);
  } catch (e) {
    logEvent_('WARN', book.label + ': ' + originalName, 'Could not move to Needs Attention: ' + e.message);
  }
  logEvent_('WARN', book.label + ': ' + originalName, `Recorded ${id} as Needs Review: ${reason}`);
}

/* ============================================================================
 * Helpers
 * ========================================================================== */

/**
 * Stores an untouched, byte-for-byte copy of the scan in
 * <Book>/Originals/YYYY/MM, named "<ID>_<original file name>".
 * Never throws: a failed copy is noted on the row instead of losing the receipt.
 * @param {File} file
 * @param {Object} bf - this book's folders
 * @param {string} id
 * @param {string} originalName
 * @param {Date} date - decides the YYYY/MM folder
 * @param {string[]} notes - a message is pushed here if the copy fails
 * @return {File|null} the copy, or null if disabled or failed
 */
function archiveOriginal_(file, bf, id, originalName, date, notes) {
  if (!settingIsYes_('Keep Original Copies')) return null;
  try {
    return file.makeCopy(id + '_' + originalName, monthFolder_(bf.originals, date));
  } catch (e) {
    notes.push('Original copy failed: ' + e.message);
    return null;
  }
}

/**
 * Builds a "View" hyperlink formula for a Drive URL.
 * @param {string} url
 * @return {string}
 */
function linkFormula_(url) {
  return '=HYPERLINK("' + String(url).replace(/"/g, '') + '","View")';
}

/**
 * Cleans Gemini output into sheet-ready values.
 * @param {Object} d - raw extraction
 * @param {{name: string, type: string}[]} categories
 * @param {Object} book
 * @return {Object}
 */
function normalizeFields_(d, categories, book) {
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
    type: hasCol_(book, 'type') ? cat.type : '',
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
 * Generates the next receipt ID for a book, e.g. W-2026-0042 or H-2026-0007.
 * Each book has its own counter, reset every year.
 * Called only while the script lock is held.
 * @param {Object} book
 * @return {string}
 */
function nextReceiptId_(book) {
  const props = PropertiesService.getScriptProperties();
  const propKey = CONFIG.PROP_COUNTER_PREFIX + book.idPrefix;
  const year = new Date().getFullYear();
  const raw = props.getProperty(propKey) || '';
  const [savedYear, savedN] = raw.split(':');
  const n = (Number(savedYear) === year ? Number(savedN) || 0 : 0) + 1;
  props.setProperty(propKey, year + ':' + n);
  return book.idPrefix + '-' + year + '-' + String(n).padStart(4, '0');
}

/**
 * Builds a descriptive file name: 2026-10-30_Shell_42.17_W-2026-0001.pdf
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
 * Returns <parent>/YYYY/MM for a date, creating folders as needed.
 * @param {Folder} parent - a book's Processed or Originals folder
 * @param {Date} date
 * @return {Folder}
 */
function monthFolder_(parent, date) {
  const tz = Session.getScriptTimeZone();
  const y = getOrCreateSubfolder_(parent, Utilities.formatDate(date, tz, 'yyyy'));
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
 * Appends a row built from a {key: value} object. Keys the book doesn't have
 * are ignored, and missing keys are left blank. Adds the Select checkbox if
 * the book has one.
 * @param {Object} book
 * @param {Sheet} sheet
 * @param {Object} obj
 */
function writeRow_(book, sheet, obj) {
  const row = book.columns.map(c => (obj[c.key] === undefined ? '' : obj[c.key]));
  sheet.appendRow(row);
  if (hasCol_(book, 'select')) {
    sheet.getRange(sheet.getLastRow(), col_(book, 'select')).insertCheckboxes();
  }
}
