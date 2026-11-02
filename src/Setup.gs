/**
 * ============================================================================
 * Setup.gs — Menu, first-time setup, folders, sheets, trigger management
 * ----------------------------------------------------------------------------
 * Created:  2026-10-30
 * Updated:  2026-11-02  v2.0 — per-book folders/tabs, Originals archive
 *           2026-11-02  v2.1 — inboxes named "Work Inbox"/"Home Inbox";
 *                       pay period settings; select-last-period menu item
 * Author:   Allen McIntosh (with Claude)
 *
 * Run "Expenses > Setup > Run first-time setup" once. It is safe to re-run:
 * existing folders, sheets, and data are kept; only missing pieces are added.
 *
 * Drive layout created under My Drive:
 *   Receipts/
 *     Work/  Work Inbox  Processed/YYYY/MM  Needs Attention  Originals/YYYY/MM
 *     Home/  Home Inbox  Processed/YYYY/MM  Needs Attention  Originals/YYYY/MM
 * ============================================================================
 */

/* ----------------------------------------------------------------------------
 * Default settings rows: Key, Value, Notes. Built in a function so it never
 * depends on .gs file load order.
 * ------------------------------------------------------------------------- */
function getDefaultSettings_() {
  return [
    ['Your Name', '', 'Used to sign reimbursement emails'],
    ['Reimbursement Email To', '', 'Who receives Work reimbursement requests (comma-separate multiple)'],
    ['Reimbursement Email CC', '', 'Optional CC (e.g., your work address)'],
    ['Send From Alias', '', 'Optional: send reimbursement email from this address. Must already be set up in Gmail > Settings > Accounts > "Send mail as". Blank = your personal Gmail'],
    ['Subject Prefix', 'Expense Reimbursement Request', 'Start of the email subject line'],
    ['Send Mode', 'Draft', 'Draft = create a Gmail draft for you to review; Send = send immediately'],
    ['Require Business Purpose', 'Yes', 'Yes = block sending if a selected row has no Business Purpose'],
    ['Default Fuel Purpose', '', 'Optional: auto-filled Business Purpose for new Work Fuel receipts'],
    ['Pay Period Anchor Start', PAY_PERIOD_DEFAULT_ANCHOR, 'Start date (a Sunday) of any one pay period. Others are counted from it'],
    ['Pay Period Length (days)', PAY_PERIOD_DEFAULT_LENGTH, '14 = biweekly'],
    ['Keep Original Copies', 'Yes', 'Yes = store an untouched copy of every scan in the book\'s Originals folder'],
    ['Gemini Model', CONFIG.DEFAULT_GEMINI_MODEL, 'Model ID used to read receipts']
  ];
}

/* ============================================================================
 * Menu
 * ========================================================================== */

/**
 * Builds the custom "Expenses" menu when the spreadsheet opens.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Expenses')
    .addItem('Process both Inboxes now', 'processInboxFromMenu')
    .addSeparator()
    .addItem('Work: Select receipts for last pay period', 'selectLastPayPeriod')
    .addItem('Work: Prepare reimbursement email (selected rows)', 'prepareReimbursement')
    .addItem('Work: Mark selected as Reimbursed', 'markSelectedReimbursed')
    .addItem('Work: Undo a reimbursement batch...', 'undoReimbursementBatch')
    .addSeparator()
    .addItem('Open Inbox folders', 'showInboxLink')
    .addSubMenu(ui.createMenu('Setup')
      .addItem('Run first-time setup', 'runSetup')
      .addItem('Set Gemini API key...', 'setApiKey')
      .addItem('Rebuild Work Summary tab', 'rebuildSummary')
      .addItem('Turn on auto-processing (every 10 min)', 'installTrigger')
      .addItem('Turn off auto-processing', 'removeTriggers'))
    .addToUi();
}

/* ============================================================================
 * First-time setup
 * ========================================================================== */

/**
 * Creates folders, sheets, validation, summary, and the processing trigger.
 * Safe to re-run. Holds the script lock so two overlapping runs (e.g., the
 * menu clicked again while the first run is still going after the
 * authorization prompt) can't both create folders or settings rows.
 */
function runSetup() {
  const ui = SpreadsheetApp.getUi();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5 * 1000)) {
    ui.alert('Setup is already running', 'Another setup or processing run is in progress. ' +
      'Wait a minute, then run setup again only if the tabs or folders are missing.', ui.ButtonSet.OK);
    return;
  }
  try {
    runSetupLocked_(ui);
  } finally {
    lock.releaseLock();
  }
}

/**
 * The body of runSetup, run while the script lock is held.
 * @param {Ui} ui
 */
function runSetupLocked_(ui) {
  const folders = ensureFolders_();
  ensureSettingsSheet_();
  getBooks_().forEach(book => {
    ensureCategoriesSheet_(book);
    ensureExpensesSheet_(book);
  });
  ensureLogSheet_();
  buildSummarySheet_();
  arrangeTabs_();
  installTrigger(true);

  const hasKey = !!PropertiesService.getScriptProperties().getProperty(CONFIG.PROP_API_KEY);
  const links = getBooks_().map(b =>
    `${b.label} Inbox: ${folders.books[b.key].inbox.getUrl()}`).join('\n');
  ui.alert(
    'Setup complete',
    'Drive folders created under "' + CONFIG.ROOT_FOLDER_NAME + '".\n' + links + '\n\n' +
    'Next steps:\n' +
    (hasKey ? '' : '1. Expenses > Setup > Set Gemini API key\n') +
    '2. Fill in the Settings tab (email recipient, your name)\n' +
    '3. Scan a receipt into the Work or Home Inbox with the Drive app\n' +
    '4. Wait up to ' + CONFIG.TRIGGER_MINUTES + ' minutes, or use Expenses > Process both Inboxes now',
    ui.ButtonSet.OK
  );
}

/**
 * Prompts for the Gemini API key and stores it in Script Properties
 * (never on a sheet, so it is not visible to anyone the sheet is shared with).
 */
function setApiKey() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('Gemini API key',
    'Paste your API key from Google AI Studio (aistudio.google.com > Get API key):',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const key = res.getResponseText().trim();
  if (!key) { ui.alert('No key entered. Nothing changed.'); return; }
  PropertiesService.getScriptProperties().setProperty(CONFIG.PROP_API_KEY, key);
  ui.alert('API key saved.');
}

/**
 * Menu wrapper (functions ending in "_" are private and can't be menu targets).
 */
function rebuildSummary() {
  buildSummarySheet_();
  SpreadsheetApp.getUi().alert('Work Summary tab rebuilt.');
}

/**
 * Shows links to both Inbox folders (handy for bookmarking on your phone).
 */
function showInboxLink() {
  const folders = ensureFolders_();
  const p = 'style="font-family:Arial;margin:6px 0"';
  const rows = getBooks_().map(b => {
    const url = folders.books[b.key].inbox.getUrl();
    return `<p ${p}><b>${escapeHtml_(b.label)}:</b> <a href="${url}" target="_blank">` +
      escapeHtml_(CONFIG.ROOT_FOLDER_NAME + ' / ' + b.folderName + ' / ' + inboxName_(b)) + '</a></p>';
  }).join('');
  const html = HtmlService.createHtmlOutput(
    `<p ${p}>Scan each receipt into the matching folder:</p>` + rows
  ).setWidth(380).setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(html, 'Receipt Inboxes');
}

/* ============================================================================
 * Folders
 * ========================================================================== */

/**
 * Returns the root folder plus each book's inbox/processed/attention/originals
 * folders, creating any that are missing. IDs are stored in Script Properties
 * so renaming or moving the folders in Drive later does not break anything.
 * @return {{root: Folder, books: Object<string, {root: Folder, inbox: Folder,
 *           processed: Folder, attention: Folder, originals: Folder}>}}
 */
function ensureFolders_() {
  const props = PropertiesService.getScriptProperties();
  const root = getOrCreateById_(props, CONFIG.PROP_ROOT_ID, () =>
    DriveApp.createFolder(CONFIG.ROOT_FOLDER_NAME));

  const out = { root: root, books: {} };
  getBooks_().forEach(book => {
    const prefix = CONFIG.PROP_FOLDER_PREFIX + book.key.toUpperCase() + '_';
    const bookRoot = getOrCreateById_(props, prefix + 'ROOT_ID', () =>
      getOrCreateSubfolder_(root, book.folderName));
    const child = (role, name) => getOrCreateById_(props, prefix + role + '_ID', () =>
      getOrCreateSubfolder_(bookRoot, name));

    // Inboxes carry the book name ("Work Inbox") so the two are easy to tell
    // apart in the Drive app's save-to picker. Older "Inbox" folders are
    // renamed in place; the stored ID keeps everything else working.
    const inbox = child('INBOX', inboxName_(book));
    if (inbox.getName() !== inboxName_(book)) inbox.setName(inboxName_(book));

    out.books[book.key] = {
      root: bookRoot,
      inbox: inbox,
      processed: child('PROCESSED', CONFIG.PROCESSED_FOLDER_NAME),
      attention: child('ATTENTION', CONFIG.ATTENTION_FOLDER_NAME),
      originals: child('ORIGINALS', CONFIG.ORIGINALS_FOLDER_NAME)
    };
  });
  return out;
}

/**
 * Display name of a book's Inbox folder, e.g. "Work Inbox".
 * @param {Object} book
 * @return {string}
 */
function inboxName_(book) {
  return book.label + ' ' + CONFIG.INBOX_FOLDER_NAME;
}

/**
 * Looks up a folder by a stored ID; if missing or trashed, creates it.
 * @param {Properties} props
 * @param {string} propKey
 * @param {function(): Folder} createFn
 * @return {Folder}
 */
function getOrCreateById_(props, propKey, createFn) {
  const id = props.getProperty(propKey);
  if (id) {
    try {
      const f = DriveApp.getFolderById(id);
      if (!f.isTrashed()) return f;
    } catch (e) { /* fall through and recreate */ }
  }
  const folder = createFn();
  props.setProperty(propKey, folder.getId());
  return folder;
}

/**
 * Returns a named subfolder of parent, creating it if needed.
 * @param {Folder} parent
 * @param {string} name
 * @return {Folder}
 */
function getOrCreateSubfolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/* ============================================================================
 * Sheets
 * ========================================================================== */

/**
 * Gets a sheet by name or creates it.
 * @param {string} name
 * @return {Sheet}
 */
function getOrCreateSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/**
 * Settings tab: Key | Value | Notes. Adds any missing keys, keeps your values.
 */
function ensureSettingsSheet_() {
  const sh = getOrCreateSheet_(CONFIG.SHEET_SETTINGS);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 3).setValues([['Key', 'Value', 'Notes']])
      .setFontWeight('bold').setBackground('#e8eaed');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 200); sh.setColumnWidth(2, 280); sh.setColumnWidth(3, 520);
  }
  removeDuplicateSettings_(sh);
  const existing = sh.getLastRow() > 1
    ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0]).trim())
    : [];
  getDefaultSettings_().forEach(row => {
    if (existing.indexOf(row[0]) < 0) sh.appendRow(row);
  });

  // Dropdowns for the Yes/No and Draft/Send settings
  const keys = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => r[0]);
  const setList = (key, list) => {
    const i = keys.indexOf(key);
    if (i >= 0) sh.getRange(i + 2, 2).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(list, true).build());
  };
  setList('Send Mode', ['Draft', 'Send']);
  setList('Require Business Purpose', ['Yes', 'No']);
  setList('Keep Original Copies', ['Yes', 'No']);
  const anchorRow = keys.indexOf('Pay Period Anchor Start');
  if (anchorRow >= 0) sh.getRange(anchorRow + 2, 2).setNumberFormat('yyyy-mm-dd');
  SETTINGS_CACHE_ = null;
}

/**
 * Deletes repeated keys on the Settings tab, keeping the first row for each
 * key. If the first row's value is blank but a later copy has one, that value
 * is moved up so nothing you typed is lost.
 * @param {Sheet} sh
 */
function removeDuplicateSettings_(sh) {
  if (sh.getLastRow() < 3) return;
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  const firstRow = {};   // key -> sheet row of first occurrence
  const toDelete = [];
  data.forEach((r, i) => {
    const key = String(r[0]).trim();
    const row = i + 2;
    if (!key) return;
    if (!(key in firstRow)) { firstRow[key] = row; return; }
    const keep = sh.getRange(firstRow[key], 2);
    if (String(keep.getValue()).trim() === '' && String(r[1]).trim() !== '') keep.setValue(r[1]);
    toDelete.push(row);
  });
  // Delete bottom-up so earlier row numbers stay valid
  toDelete.reverse().forEach(row => sh.deleteRow(row));
  if (toDelete.length) logEvent_('INFO', 'Setup', `Removed ${toDelete.length} duplicate Settings row(s).`);
}

/**
 * A book's Categories tab. Only seeds defaults when the tab is empty, so your
 * edits are preserved.
 * @param {Object} book
 */
function ensureCategoriesSheet_(book) {
  const sh = getOrCreateSheet_(book.categoriesSheet);
  const cols = book.categoryColumns;
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, cols.length).setValues([cols.map(c => c.header)])
      .setFontWeight('bold').setBackground('#e8eaed');
    sh.getRange(2, 1, book.defaultCategories.length, cols.length).setValues(book.defaultCategories);
    sh.setFrozenRows(1);
    cols.forEach((c, i) => sh.setColumnWidth(i + 1, c.width));
  }
  if (catCol_(book, 'type')) {
    const L = colLetter_(catCol_(book, 'type'));
    sh.getRange(L + '2:' + L).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList([TYPES.REIMBURSABLE, TYPES.TRACK_ONLY], true).build());
  }
  if (catCol_(book, 'budget')) {
    const L = colLetter_(catCol_(book, 'budget'));
    sh.getRange(L + '2:' + L).setNumberFormat('$#,##0.00');
  }
}

/**
 * A book's Expenses tab: headers, widths, formats, dropdowns, filter.
 * @param {Object} book
 */
function ensureExpensesSheet_(book) {
  const sh = getOrCreateSheet_(book.expensesSheet);
  const n = book.columns.length;

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, n).setValues([book.columns.map(c => c.header)])
      .setFontWeight('bold').setBackground(book.headerColor).setFontColor('#ffffff')
      .setWrap(true).setVerticalAlignment('middle');
    sh.setFrozenRows(1);
    sh.setFrozenColumns(col_(book, 'vendor'));
    book.columns.forEach((c, i) => {
      sh.setColumnWidth(i + 1, c.width);
      if (c.hidden) sh.hideColumns(i + 1);
    });
  }

  const range = k => {
    const L = colLetter_(col_(book, k));
    return sh.getRange(L + '2:' + L);
  };
  const ifCol = (k, fn) => { if (hasCol_(book, k)) fn(range(k)); };

  // Number / date formats
  range('added').setNumberFormat('yyyy-mm-dd hh:mm');
  range('date').setNumberFormat('yyyy-mm-dd');
  ['subtotal', 'tax', 'total'].forEach(k => range(k).setNumberFormat('$#,##0.00'));
  range('gallons').setNumberFormat('0.000');
  range('ppg').setNumberFormat('$0.000');
  range('payment').setNumberFormat('@'); // keep leading zeros on card last-4
  ifCol('submitted', r => r.setNumberFormat('yyyy-mm-dd'));

  // Dropdowns
  const catSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(book.categoriesSheet);
  range('category').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(catSheet.getRange('A2:A'), true).setAllowInvalid(true).build());
  ifCol('type', r => r.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList([TYPES.REIMBURSABLE, TYPES.TRACK_ONLY], true).build()));
  range('status').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(book.statuses, true).build());

  // Status colors
  const statusRange = range('status');
  sh.setConditionalFormatRules(book.statuses.map(s =>
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(s)
      .setBackground(STATUS_COLORS[s] || '#ffffff').setRanges([statusRange]).build()));

  // Filter over the full sheet height so newly added rows are included
  if (!sh.getFilter()) sh.getRange(1, 1, sh.getMaxRows(), n).createFilter();
}

/**
 * Log tab for processing history and errors (shared by both books).
 */
function ensureLogSheet_() {
  const sh = getOrCreateSheet_(CONFIG.SHEET_LOG);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 4).setValues([['Timestamp', 'Level', 'Book: File / Context', 'Message']])
      .setFontWeight('bold').setBackground('#e8eaed');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 150); sh.setColumnWidth(3, 300); sh.setColumnWidth(4, 520);
    sh.getRange('A2:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  }
}

/**
 * Puts tabs in a sensible order and removes the blank starter "Sheet1".
 */
function arrangeTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const work = getBook_('work'), home = getBook_('home');
  const order = [work.expensesSheet, home.expensesSheet, CONFIG.SHEET_SUMMARY,
                 work.categoriesSheet, home.categoriesSheet, CONFIG.SHEET_SETTINGS, CONFIG.SHEET_LOG];
  order.forEach((name, i) => {
    const sh = ss.getSheetByName(name);
    if (sh) { ss.setActiveSheet(sh); ss.moveActiveSheet(i + 1); }
  });
  const starter = ss.getSheetByName('Sheet1');
  if (starter && starter.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(starter);
  ss.setActiveSheet(ss.getSheetByName(work.expensesSheet));
}

/* ============================================================================
 * Work Summary tab (budget tracking, Work book only)
 * ========================================================================== */

/**
 * (Re)builds the Work Summary tab: monthly budget vs. actual per category,
 * reimbursement totals, and a month-by-category history table.
 * Re-run after adding Work categories.
 */
function buildSummarySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const book = getBook_('work');
  const sh = getOrCreateSheet_(CONFIG.SHEET_SUMMARY);
  sh.clear();
  sh.getRange('A:Z').clearDataValidations();

  const E = "'" + book.expensesSheet + "'!";
  const L = k => colLetter_(col_(book, k));
  const colRef = k => E + L(k) + ':' + L(k);
  const total = colRef('total'), date = colRef('date'), cat = colRef('category');
  const type = colRef('type'), status = colRef('status');

  // Month picker
  sh.getRange('A1').setValue('Month:').setFontWeight('bold');
  sh.getRange('B1').setFormula('=EOMONTH(TODAY(),-1)+1').setNumberFormat('mmmm yyyy')
    .setNote('Type any date in this month to view it, e.g. 9/1/2026. ' +
             'Restore the formula =EOMONTH(TODAY(),-1)+1 to follow the current month.');
  const inMonth = `,${date},">="&$B$1,${date},"<"&EDATE($B$1,1)`;

  // Reimbursement snapshot
  sh.getRange('D1').setValue('Reimbursement').setFontWeight('bold');
  sh.getRange('D2:E4').setValues([
    ['Unsubmitted', `=SUMIFS(${total},${type},"${TYPES.REIMBURSABLE}",${status},"${STATUS.UNSUBMITTED}")`],
    ['Submitted, awaiting payment', `=SUMIFS(${total},${status},"${STATUS.SUBMITTED}")`],
    ['Reimbursed (this month\'s receipts)', `=SUMIFS(${total},${status},"${STATUS.REIMBURSED}"${inMonth})`]
  ]);
  sh.getRange('E2:E4').setNumberFormat('$#,##0.00');

  // Budget vs actual for the selected month
  const startRow = 6;
  sh.getRange(startRow, 1, 1, 5)
    .setValues([['Category', 'Monthly Budget', 'Spent', 'Remaining', '% Used']])
    .setFontWeight('bold').setBackground(book.headerColor).setFontColor('#ffffff');

  const catSheet = ss.getSheetByName(book.categoriesSheet);
  const cats = catSheet.getLastRow() > 1
    ? catSheet.getRange(2, 1, catSheet.getLastRow() - 1, 1).getValues().map(r => r[0]).filter(String)
    : [];

  // Budget column position on the Categories tab, looked up by key
  const budgetIdx = catCol_(book, 'budget');
  const C = "'" + book.categoriesSheet + "'!A:" + colLetter_(budgetIdx);
  const rows = cats.map((c, i) => {
    const r = startRow + 1 + i;
    // Look the budget up by name so reordering the Categories tab is harmless
    const look = `IFERROR(VLOOKUP(A${r},${C},${budgetIdx},FALSE),"")`;
    return [
      c,
      `=IF(${look}="","",${look})`,
      `=SUMIFS(${total},${cat},A${r}${inMonth})`,
      `=IF(B${r}="","",B${r}-C${r})`,
      `=IF(OR(B${r}="",B${r}=0),"",C${r}/B${r})`
    ];
  });
  if (rows.length) sh.getRange(startRow + 1, 1, rows.length, 5).setValues(rows);

  const totalRow = startRow + 1 + rows.length;
  const first = startRow + 1, last = totalRow - 1;
  sh.getRange(totalRow, 1, 1, 5).setValues([[
    'TOTAL',
    rows.length ? `=SUM(B${first}:B${last})` : 0,
    rows.length ? `=SUM(C${first}:C${last})` : 0,
    `=IF(B${totalRow}=0,"",B${totalRow}-C${totalRow})`,
    `=IF(B${totalRow}=0,"",C${totalRow}/B${totalRow})`
  ]]).setFontWeight('bold').setBorder(true, null, null, null, null, null);

  sh.getRange(first, 2, rows.length + 1, 3).setNumberFormat('$#,##0.00');
  sh.getRange(first, 5, rows.length + 1, 1).setNumberFormat('0%');

  // Over-budget highlight
  const pctRange = sh.getRange(first, 5, rows.length + 1, 1);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(1)
      .setBackground('#f4cccc').setRanges([pctRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0.85, 1)
      .setBackground('#fff2cc').setRanges([pctRange]).build()
  ]);

  // Monthly history: one row per month, one column per category
  const histRow = totalRow + 3;
  sh.getRange(histRow - 1, 1).setValue('Monthly history (all months)').setFontWeight('bold');
  const D = L('date'), F = L('category'), J = L('total');
  const lastLetter = colLetter_(book.columns.length);
  sh.getRange(histRow, 1).setFormula(
    `=IFERROR(QUERY(${E}A1:${lastLetter}, ` +
    `"select year(${D}), month(${D})+1, sum(${J}) where ${D} is not null and ${J} is not null ` +
    `group by year(${D}), month(${D})+1 pivot ${F} ` +
    `label year(${D}) 'Year', month(${D})+1 'Month'", 1), "No data yet")`
  );

  sh.setColumnWidth(1, 210); sh.setColumnWidth(2, 120); sh.setColumnWidth(3, 110);
  sh.setColumnWidth(4, 230); sh.setColumnWidth(5, 110);
  sh.setFrozenRows(1);
}

/* ============================================================================
 * Triggers
 * ========================================================================== */

/**
 * Installs the time-driven trigger that processes both Inboxes.
 * @param {boolean=} silent - skip the confirmation alert (used by setup)
 */
function installTrigger(silent) {
  removeTriggers(true);
  ScriptApp.newTrigger('processInbox').timeBased().everyMinutes(CONFIG.TRIGGER_MINUTES).create();
  if (silent !== true) {
    SpreadsheetApp.getUi().alert('Auto-processing is on (every ' + CONFIG.TRIGGER_MINUTES + ' minutes).');
  }
}

/**
 * Removes this project's processInbox triggers.
 * @param {boolean=} silent
 */
function removeTriggers(silent) {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'processInbox')
    .forEach(t => ScriptApp.deleteTrigger(t));
  if (silent !== true) SpreadsheetApp.getUi().alert('Auto-processing is off.');
}
