/**
 * ============================================================================
 * Config.gs — Receipt Tracker constants and shared helpers
 * ----------------------------------------------------------------------------
 * Created:  2026-10-30
 * Updated:  2026-11-02  v2.0 — Work/Home books, original-copy archive
 * Author:   Allen McIntosh (with Claude)
 *
 * Generic constants and utilities used by every file. The Work and Home
 * "books" (their columns, categories, folders, and statuses) are defined in
 * Books.gs. Values you are likely to tweak (email recipient, send mode, model
 * name, default purpose) live on the "Settings" sheet so you never have to
 * edit code for them.
 * ============================================================================
 */

const CONFIG = {
  // Top-level Drive folder (created under My Drive). Each book gets its own
  // subfolder with Inbox / Processed / Needs Attention / Originals inside.
  ROOT_FOLDER_NAME: 'Receipts',
  INBOX_FOLDER_NAME: 'Inbox',
  PROCESSED_FOLDER_NAME: 'Processed',
  ATTENTION_FOLDER_NAME: 'Needs Attention',
  ORIGINALS_FOLDER_NAME: 'Originals',

  // Shared sheet (tab) names. Per-book tabs are named in Books.gs.
  SHEET_SUMMARY: 'Work Summary',
  SHEET_SETTINGS: 'Settings',
  SHEET_LOG: 'Log',

  // Gemini. "gemini-flash-latest" is Google's rolling alias for the current
  // Flash model, so it won't break when versions retire. Override on Settings.
  DEFAULT_GEMINI_MODEL: 'gemini-flash-latest',
  GEMINI_ENDPOINT: 'https://generativelanguage.googleapis.com/v1beta/models/',
  GEMINI_MAX_ATTEMPTS: 3,

  // Limits
  MAX_FILE_BYTES: 15 * 1024 * 1024,        // inline Gemini requests cap ~20 MB
  MAX_EMAIL_ATTACH_BYTES: 24 * 1024 * 1024, // Gmail cap is 25 MB
  MAX_RUNTIME_MS: 4.5 * 60 * 1000,          // stay under Apps Script 6-min limit
  TRIGGER_MINUTES: 10,

  // File types Gemini can read directly
  SUPPORTED_MIME_TYPES: [
    'application/pdf', 'image/jpeg', 'image/png',
    'image/webp', 'image/heic', 'image/heif'
  ],

  // Script Properties keys. Per-book folder IDs and counters are built from
  // these prefixes plus the book key, e.g. FOLDER_WORK_INBOX_ID, COUNTER_W.
  PROP_API_KEY: 'GEMINI_API_KEY',
  PROP_ROOT_ID: 'FOLDER_ROOT_ID',
  PROP_FOLDER_PREFIX: 'FOLDER_',
  PROP_COUNTER_PREFIX: 'COUNTER_'
};

/* ----------------------------------------------------------------------------
 * Expense types and statuses (used for dropdowns and logic)
 * ------------------------------------------------------------------------- */
const TYPES = {
  REIMBURSABLE: 'Reimbursable',
  TRACK_ONLY: 'Track Only'
};

const STATUS = {
  NEEDS_REVIEW: 'Needs Review',  // extraction failed or looked shaky (both books)
  UNSUBMITTED: 'Unsubmitted',    // Work: reimbursable, not yet sent
  SUBMITTED: 'Submitted',        // Work: included in a reimbursement email
  REIMBURSED: 'Reimbursed',      // Work: money received
  REJECTED: 'Rejected',          // Work: reimbursement denied
  TRACKED: 'Tracked',            // Work: track-only expense, nothing to do
  RECORDED: 'Recorded'           // Home: stored and read cleanly
};

/** Background colors for each status (conditional formatting). */
const STATUS_COLORS = {
  'Needs Review': '#f4cccc',
  'Unsubmitted': '#fff2cc',
  'Submitted': '#cfe2f3',
  'Reimbursed': '#d9ead3',
  'Rejected': '#e6b8af',
  'Tracked': '#eeeeee',
  'Recorded': '#eeeeee'
};

/* ============================================================================
 * Column lookup (per book)
 * ========================================================================== */

/**
 * Returns the 1-based column number for a column key in a book's sheet.
 * @param {Object} book - from getBook_()
 * @param {string} key  - key from book.columns
 * @return {number}
 */
function col_(book, key) {
  if (!book.colIndex_) {
    book.colIndex_ = {};
    book.columns.forEach((c, i) => { book.colIndex_[c.key] = i + 1; });
  }
  const n = book.colIndex_[key];
  if (!n) throw new Error('Unknown column "' + key + '" in ' + book.label + ' book');
  return n;
}

/**
 * True if the book's sheet has a column with this key.
 * @param {Object} book
 * @param {string} key
 * @return {boolean}
 */
function hasCol_(book, key) {
  return book.columns.some(c => c.key === key);
}

/**
 * Returns the 1-based column number for a key on a book's Categories tab,
 * or 0 if that book's Categories tab has no such column.
 * @param {Object} book
 * @param {string} key - key from book.categoryColumns
 * @return {number}
 */
function catCol_(book, key) {
  return book.categoryColumns.findIndex(c => c.key === key) + 1;
}

/**
 * Converts a column number to its A1 letter (1 -> A, 27 -> AA).
 * @param {number} n
 * @return {string}
 */
function colLetter_(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/**
 * Turns a sheet row array into a {key: value} object for a book.
 * @param {Object} book
 * @param {Array} row
 * @return {Object}
 */
function rowToObject_(book, row) {
  const o = {};
  book.columns.forEach((c, i) => { o[c.key] = row[i]; });
  return o;
}

/* ============================================================================
 * Settings sheet access
 * ========================================================================== */

let SETTINGS_CACHE_ = null;

/**
 * Reads a value from the Settings sheet (Key | Value | Notes).
 * Cached for the life of one execution.
 * @param {string} key
 * @return {string} trimmed value, or '' if missing
 */
function getSetting_(key) {
  if (!SETTINGS_CACHE_) {
    SETTINGS_CACHE_ = {};
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_SETTINGS);
    if (sh && sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(r => {
        if (r[0]) SETTINGS_CACHE_[String(r[0]).trim()] = String(r[1]).trim();
      });
    }
  }
  return SETTINGS_CACHE_[key] || '';
}

/**
 * Reads a Yes/No setting. Anything other than "No" counts as Yes, so a
 * missing or mistyped value falls back to the safer default.
 * @param {string} key
 * @return {boolean}
 */
function settingIsYes_(key) {
  return getSetting_(key).toLowerCase() !== 'no';
}

/* ============================================================================
 * Small shared utilities
 * ========================================================================== */

/**
 * Prevents spreadsheet formula injection from text pulled off a receipt
 * (a vendor name starting with "=" would otherwise execute as a formula).
 * @param {*} v
 * @return {*}
 */
function safeText_(v) {
  if (v === null || v === undefined) return '';
  const s = String(v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

/**
 * Escapes text for safe inclusion in an HTML email body.
 * @param {*} v
 * @return {string}
 */
function escapeHtml_(v) {
  return String(v === null || v === undefined ? '' : v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/**
 * Formats a number as US currency text, e.g. 1234.5 -> "$1,234.50".
 * @param {number} n
 * @return {string}
 */
function money_(n) {
  const v = Number(n) || 0;
  return '$' + v.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

/**
 * Writes a line to the Log sheet (and Logger). Never throws.
 * @param {string} level   - INFO | WARN | ERROR
 * @param {string} context - book + file name, or other context
 * @param {string} message
 */
function logEvent_(level, context, message) {
  try {
    Logger.log(`[${level}] ${context}: ${message}`);
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_LOG);
    if (sh) sh.appendRow([new Date(), level, safeText_(context), safeText_(message)]);
  } catch (e) {
    Logger.log('Log write failed: ' + e);
  }
}
