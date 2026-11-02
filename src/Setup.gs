/**
 * ============================================================================
 * Setup.gs — Menu, first-time setup, folders, sheets, trigger management
 * ----------------------------------------------------------------------------
 * Created:  2026-10-30
 * Author:   Allen McIntosh (with Claude)
 *
 * Run "Expenses > Setup > Run first-time setup" once. It is safe to re-run:
 * existing folders, sheets, and data are kept; only missing pieces are added.
 * ============================================================================
 */

/* ----------------------------------------------------------------------------
 * Default categories. "Default Type" decides whether a new receipt in that
 * category starts as Reimbursable or Track Only (you can override per row).
 * The description is sent to Gemini to help it pick the right category.
 * ------------------------------------------------------------------------- */
const DEFAULT_CATEGORIES = [
  // Category, Default Type, Monthly Budget, Description (sent to AI)
  ['Fuel', TYPES.REIMBURSABLE, '', 'Gasoline/diesel purchases at a fuel station'],
  ['Parking & Tolls', TYPES.REIMBURSABLE, '', 'Parking garages, meters, toll charges'],
  ['Meals', TYPES.REIMBURSABLE, '', 'Restaurants, food purchased while traveling for work'],
  ['Travel - Lodging', TYPES.REIMBURSABLE, '', 'Hotels and lodging'],
  ['IT Equipment', TYPES.TRACK_ONLY, '', 'Computers, monitors, network gear, printers, phones, tablets'],
  ['Cables & Accessories', TYPES.TRACK_ONLY, '', 'Cables, adapters, mice, keyboards, mounts, small peripherals'],
  ['Software & Subscriptions', TYPES.TRACK_ONLY, '', 'Software licenses, SaaS, app subscriptions'],
  ['Office Supplies', TYPES.TRACK_ONLY, '', 'Paper, toner, labels, general office items'],
  ['Tools & Maintenance', TYPES.TRACK_ONLY, '', 'Hand tools, hardware store items, repair parts'],
  ['Shipping', TYPES.TRACK_ONLY, '', 'Postage, courier, shipping fees'],
  ['Other', TYPES.TRACK_ONLY, '', 'Anything that does not fit another category']
];

/* ----------------------------------------------------------------------------
 * Default settings rows: Key, Value, Notes
 * ------------------------------------------------------------------------- */
const DEFAULT_SETTINGS = [
  ['Your Name', '', 'Used to sign reimbursement emails'],
  ['Reimbursement Email To', '', 'Who receives reimbursement requests (comma-separate multiple)'],
  ['Reimbursement Email CC', '', 'Optional CC (e.g., your own address)'],
  ['Subject Prefix', 'Expense Reimbursement Request', 'Start of the email subject line'],
  ['Send Mode', 'Draft', 'Draft = create a Gmail draft for you to review; Send = send immediately'],
  ['Require Business Purpose', 'Yes', 'Yes = block sending if a selected row has no Business Purpose'],
  ['Default Fuel Purpose', '', 'Optional: auto-filled Business Purpose for new Fuel receipts'],
  ['Gemini Model', CONFIG.DEFAULT_GEMINI_MODEL, 'Model ID used to read receipts']
];

/* ============================================================================
 * Menu
 * ========================================================================== */

/**
 * Builds the custom "Expenses" menu when the spreadsheet opens.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Expenses')
    .addItem('Process Inbox now', 'processInboxFromMenu')
    .addSeparator()
    .addItem('Prepare reimbursement email (selected rows)', 'prepareReimbursement')
    .addItem('Mark selected as Reimbursed', 'markSelectedReimbursed')
    .addItem('Undo a reimbursement batch...', 'undoReimbursementBatch')
    .addSeparator()
    .addItem('Open Inbox folder', 'showInboxLink')
    .addSubMenu(ui.createMenu('Setup')
      .addItem('Run first-time setup', 'runSetup')
      .addItem('Set Gemini API key...', 'setApiKey')
      .addItem('Rebuild Summary tab', 'rebuildSummary')
      .addItem('Turn on auto-processing (every 10 min)', 'installTrigger')
      .addItem('Turn off auto-processing', 'removeTriggers'))
    .addToUi();
}

/* ============================================================================
 * First-time setup
 * ========================================================================== */

/**
 * Creates folders, sheets, validation, summary, and the processing trigger.
 * Safe to re-run.
 */
function runSetup() {
  const ui = SpreadsheetApp.getUi();
  const folders = ensureFolders_();
  ensureSettingsSheet_();
  ensureCategoriesSheet_();
  ensureExpensesSheet_();
  ensureLogSheet_();
  buildSummarySheet_();
  installTrigger(true);

  const hasKey = !!PropertiesService.getScriptProperties().getProperty(CONFIG.PROP_API_KEY);
  ui.alert(
    'Setup complete',
    'Drive folder created: "' + CONFIG.ROOT_FOLDER_NAME + '/' + CONFIG.INBOX_FOLDER_NAME + '"\n' +
    folders.inbox.getUrl() + '\n\n' +
    'Next steps:\n' +
    (hasKey ? '' : '1. Expenses > Setup > Set Gemini API key\n') +
    '2. Fill in the Settings tab (email recipient, your name)\n' +
    '3. Scan a receipt into the Inbox folder with the Drive app\n' +
    '4. Wait up to ' + CONFIG.TRIGGER_MINUTES + ' minutes, or use Expenses > Process Inbox now',
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
  SpreadsheetApp.getUi().alert('Summary tab rebuilt.');
}

/**
 * Shows a link to the Inbox folder (handy for bookmarking on your phone).
 */
function showInboxLink() {
  const url = ensureFolders_().inbox.getUrl();
  const html = HtmlService.createHtmlOutput(
    '<p style="font-family:Arial">Scan receipts into this folder:</p>' +
    '<p style="font-family:Arial"><a href="' + url + '" target="_blank">' +
    escapeHtml_(CONFIG.ROOT_FOLDER_NAME + ' / ' + CONFIG.INBOX_FOLDER_NAME) + '</a></p>'
  ).setWidth(360).setHeight(120);
  SpreadsheetApp.getUi().showModalDialog(html, 'Receipt Inbox');
}

/* ============================================================================
 * Folders
 * ========================================================================== */

/**
 * Returns the root/inbox/processed/attention folders, creating any that are
 * missing. IDs are stored in Script Properties so renaming/moving the folders
 * in Drive later does not break anything.
 * @return {{root: Folder, inbox: Folder, processed: Folder, attention: Folder}}
 */
function ensureFolders_() {
  const props = PropertiesService.getScriptProperties();

  const root = getOrCreateById_(props, CONFIG.PROP_ROOT_ID, () =>
    DriveApp.createFolder(CONFIG.ROOT_FOLDER_NAME));

  const child = (propKey, name) => getOrCreateById_(props, propKey, () =>
    getOrCreateSubfolder_(root, name));

  return {
    root: root,
    inbox: child(CONFIG.PROP_INBOX_ID, CONFIG.INBOX_FOLDER_NAME),
    processed: child(CONFIG.PROP_PROCESSED_ID, CONFIG.PROCESSED_FOLDER_NAME),
    attention: child(CONFIG.PROP_ATTENTION_ID, CONFIG.ATTENTION_FOLDER_NAME)
  };
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
    sh.setColumnWidth(1, 200); sh.setColumnWidth(2, 280); sh.setColumnWidth(3, 420);
  }
  const existing = sh.getLastRow() > 1
    ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0]).trim())
    : [];
  DEFAULT_SETTINGS.forEach(row => {
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
  SETTINGS_CACHE_ = null;
}

/**
 * Categories tab: Category | Default Type | Monthly Budget | Description.
 * Only seeds defaults when the tab is empty, so your edits are preserved.
 */
function ensureCategoriesSheet_() {
  const sh = getOrCreateSheet_(CONFIG.SHEET_CATEGORIES);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 4)
      .setValues([['Category', 'Default Type', 'Monthly Budget', 'Description (helps the AI choose)']])
      .setFontWeight('bold').setBackground('#e8eaed');
    sh.getRange(2, 1, DEFAULT_CATEGORIES.length, 4).setValues(DEFAULT_CATEGORIES);
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 190); sh.setColumnWidth(2, 120);
    sh.setColumnWidth(3, 120); sh.setColumnWidth(4, 420);
  }
  sh.getRange('B2:B').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList([TYPES.REIMBURSABLE, TYPES.TRACK_ONLY], true).build());
  sh.getRange('C2:C').setNumberFormat('$#,##0.00');
}

/**
 * Expenses tab: headers, widths, formats, dropdowns, filter.
 */
function ensureExpensesSheet_() {
  const sh = getOrCreateSheet_(CONFIG.SHEET_EXPENSES);
  const n = COLUMNS.length;

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, n).setValues([COLUMNS.map(c => c.header)])
      .setFontWeight('bold').setBackground('#1f3864').setFontColor('#ffffff')
      .setWrap(true).setVerticalAlignment('middle');
    sh.setFrozenRows(1);
    sh.setFrozenColumns(col_('vendor'));
    COLUMNS.forEach((c, i) => sh.setColumnWidth(i + 1, c.width));
    sh.hideColumns(col_('fileId'));
  }

  const L = k => colLetter_(col_(k));
  const range = k => sh.getRange(L(k) + '2:' + L(k));

  // Number / date formats
  range('added').setNumberFormat('yyyy-mm-dd hh:mm');
  range('date').setNumberFormat('yyyy-mm-dd');
  range('submitted').setNumberFormat('yyyy-mm-dd');
  ['subtotal', 'tax', 'total'].forEach(k => range(k).setNumberFormat('$#,##0.00'));
  range('gallons').setNumberFormat('0.000');
  range('ppg').setNumberFormat('$0.000');
  range('payment').setNumberFormat('@'); // keep leading zeros on card last-4

  // Dropdowns
  const catSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_CATEGORIES);
  range('category').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(catSheet.getRange('A2:A'), true).setAllowInvalid(true).build());
  range('type').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList([TYPES.REIMBURSABLE, TYPES.TRACK_ONLY], true).build());
  range('status').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.keys(STATUS).map(k => STATUS[k]), true).build());

  // Status colors
  const statusRange = range('status');
  const rule = (text, bg) => SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(text).setBackground(bg).setRanges([statusRange]).build();
  sh.setConditionalFormatRules([
    rule(STATUS.NEEDS_REVIEW, '#f4cccc'),
    rule(STATUS.UNSUBMITTED, '#fff2cc'),
    rule(STATUS.SUBMITTED, '#cfe2f3'),
    rule(STATUS.REIMBURSED, '#d9ead3'),
    rule(STATUS.REJECTED, '#e6b8af'),
    rule(STATUS.TRACKED, '#eeeeee')
  ]);

  // Filter over the full sheet height so newly added rows are included
  if (!sh.getFilter()) sh.getRange(1, 1, sh.getMaxRows(), n).createFilter();
}

/**
 * Log tab for processing history and errors.
 */
function ensureLogSheet_() {
  const sh = getOrCreateSheet_(CONFIG.SHEET_LOG);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 4).setValues([['Timestamp', 'Level', 'File / Context', 'Message']])
      .setFontWeight('bold').setBackground('#e8eaed');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 150); sh.setColumnWidth(3, 280); sh.setColumnWidth(4, 520);
    sh.getRange('A2:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  }
}

/* ============================================================================
 * Summary tab (budget tracking)
 * ========================================================================== */

/**
 * (Re)builds the Summary tab: monthly budget vs. actual per category,
 * reimbursement totals, and a month-by-category history table.
 * Re-run after adding categories.
 */
function buildSummarySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getOrCreateSheet_(CONFIG.SHEET_SUMMARY);
  sh.clear();
  sh.getRange('A:Z').clearDataValidations();

  const E = "'" + CONFIG.SHEET_EXPENSES + "'!";
  const L = k => colLetter_(col_(k));
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
    .setFontWeight('bold').setBackground('#1f3864').setFontColor('#ffffff');

  const catSheet = ss.getSheetByName(CONFIG.SHEET_CATEGORIES);
  const cats = catSheet.getLastRow() > 1
    ? catSheet.getRange(2, 1, catSheet.getLastRow() - 1, 1).getValues().map(r => r[0]).filter(String)
    : [];

  const C = "'" + CONFIG.SHEET_CATEGORIES + "'!A:C";
  const rows = cats.map((c, i) => {
    const r = startRow + 1 + i;
    // Look the budget up by name so reordering the Categories tab is harmless
    const look = `IFERROR(VLOOKUP(A${r},${C},3,FALSE),"")`;
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
  const lastLetter = colLetter_(COLUMNS.length);
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
 * Installs the time-driven trigger that processes the Inbox.
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
