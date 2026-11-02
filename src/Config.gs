/**
 * ============================================================================
 * Config.gs — Receipt Tracker constants, column map, and shared helpers
 * ----------------------------------------------------------------------------
 * Created:  2026-10-30
 * Author:   Allen McIntosh (with Claude)
 *
 * Everything that rarely changes lives here. Things you are likely to tweak
 * (email recipient, send mode, model name, default purpose) live on the
 * "Settings" sheet instead so you never have to edit code for them.
 * ============================================================================
 */

const CONFIG = {
  // Drive folder names (created by Setup under My Drive)
  ROOT_FOLDER_NAME: 'Expense Receipts',
  INBOX_FOLDER_NAME: 'Inbox',
  PROCESSED_FOLDER_NAME: 'Processed',
  ATTENTION_FOLDER_NAME: 'Needs Attention',

  // Sheet (tab) names
  SHEET_EXPENSES: 'Expenses',
  SHEET_CATEGORIES: 'Categories',
  SHEET_SUMMARY: 'Summary',
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

  // Script Properties keys
  PROP_API_KEY: 'GEMINI_API_KEY',
  PROP_ROOT_ID: 'FOLDER_ROOT_ID',
  PROP_INBOX_ID: 'FOLDER_INBOX_ID',
  PROP_PROCESSED_ID: 'FOLDER_PROCESSED_ID',
  PROP_ATTENTION_ID: 'FOLDER_ATTENTION_ID',
  PROP_COUNTER: 'RECEIPT_COUNTER'
};

/* ----------------------------------------------------------------------------
 * Expense types and statuses (used for dropdowns and logic)
 * ------------------------------------------------------------------------- */
const TYPES = {
  REIMBURSABLE: 'Reimbursable',
  TRACK_ONLY: 'Track Only'
};

const STATUS = {
  NEEDS_REVIEW: 'Needs Review',  // extraction failed or looked shaky
  UNSUBMITTED: 'Unsubmitted',    // reimbursable, not yet sent
  SUBMITTED: 'Submitted',        // included in a reimbursement email
  REIMBURSED: 'Reimbursed',      // money received
  REJECTED: 'Rejected',          // reimbursement denied
  TRACKED: 'Tracked'             // track-only expense, nothing to do
};

/* ----------------------------------------------------------------------------
 * Expenses sheet column map. Order here = column order on the sheet.
 * Code always looks columns up by key, so you can reorder this list (before
 * first setup) without touching any other file.
 * ------------------------------------------------------------------------- */
const COLUMNS = [
  { key: 'select',    header: 'Select',           width: 60 },
  { key: 'id',        header: 'Receipt ID',       width: 105 },
  { key: 'added',     header: 'Added',            width: 130 },
  { key: 'date',      header: 'Receipt Date',     width: 100 },
  { key: 'vendor',    header: 'Vendor',           width: 180 },
  { key: 'category',  header: 'Category',         width: 150 },
  { key: 'type',      header: 'Type',             width: 110 },
  { key: 'subtotal',  header: 'Subtotal',         width: 90 },
  { key: 'tax',       header: 'Tax',              width: 75 },
  { key: 'total',     header: 'Total',            width: 90 },
  { key: 'payment',   header: 'Card (last 4)',    width: 90 },
  { key: 'gallons',   header: 'Gallons',          width: 75 },
  { key: 'ppg',       header: 'Price / Gal',      width: 85 },
  { key: 'purpose',   header: 'Business Purpose', width: 220 },
  { key: 'site',      header: 'Site / Project',   width: 130 },
  { key: 'status',    header: 'Status',           width: 115 },
  { key: 'submitted', header: 'Submitted On',     width: 100 },
  { key: 'batch',     header: 'Batch ID',         width: 135 },
  { key: 'receipt',   header: 'Receipt',          width: 70 },
  { key: 'items',     header: 'Items / Summary',  width: 260 },
  { key: 'notes',     header: 'Extraction Notes', width: 260 },
  { key: 'fileId',    header: 'File ID',          width: 120 }
];

/**
 * Returns the 1-based column number for a column key.
 * @param {string} key - key from COLUMNS
 * @return {number}
 */
function col_(key) {
  const idx = COLUMNS.findIndex(c => c.key === key);
  if (idx < 0) throw new Error('Unknown column key: ' + key);
  return idx + 1;
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
 * @param {string} level - INFO | WARN | ERROR
 * @param {string} file  - file name or context
 * @param {string} message
 */
function logEvent_(level, file, message) {
  try {
    Logger.log(`[${level}] ${file}: ${message}`);
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_LOG);
    if (sh) sh.appendRow([new Date(), level, safeText_(file), safeText_(message)]);
  } catch (e) {
    Logger.log('Log write failed: ' + e);
  }
}
