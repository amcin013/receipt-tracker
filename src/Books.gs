/**
 * ============================================================================
 * Books.gs — Work and Home "book" definitions
 * ----------------------------------------------------------------------------
 * Created:  2026-11-02
 * Author:   Allen McIntosh (with Claude)
 *
 * A "book" is one fully separate set of receipts: its own Drive folders,
 * Inbox, expenses tab, categories tab, ID series, and statuses. Which book a
 * receipt belongs to is decided only by the Inbox folder it was scanned into;
 * nothing is guessed.
 *
 *   Work  - reimbursement workflow + monthly budget summary (the v1 feature set)
 *   Home  - storage and data only, kept tidy so a home budget sheet can pull
 *           from it later
 *
 * Definitions are built inside a function (not as top-level constants) so
 * they never depend on the order Apps Script loads the .gs files.
 * ============================================================================
 */

let BOOKS_CACHE_ = null;

/**
 * Returns every book definition, Work first.
 * @return {Object[]}
 */
function getBooks_() {
  if (BOOKS_CACHE_) return BOOKS_CACHE_;

  /* --------------------------------------------------------------------------
   * Columns shared by both books, in the order they appear. Each book adds its
   * own columns around these. Code always looks columns up by key, so reorder
   * only BEFORE first setup; existing rows are never rearranged.
   * ----------------------------------------------------------------------- */
  const receiptCols = [
    { key: 'id',        header: 'Receipt ID',       width: 105 },
    { key: 'added',     header: 'Added',            width: 130 },
    { key: 'date',      header: 'Receipt Date',     width: 100 },
    { key: 'vendor',    header: 'Vendor',           width: 180 },
    { key: 'category',  header: 'Category',         width: 150 }
  ];
  const amountCols = [
    { key: 'subtotal',  header: 'Subtotal',         width: 90 },
    { key: 'tax',       header: 'Tax',              width: 75 },
    { key: 'total',     header: 'Total',            width: 90 },
    { key: 'payment',   header: 'Card (last 4)',    width: 90 },
    { key: 'gallons',   header: 'Gallons',          width: 75 },
    { key: 'ppg',       header: 'Price / Gal',      width: 85 }
  ];
  const fileCols = [
    { key: 'receipt',    header: 'Receipt',          width: 70 },
    { key: 'original',   header: 'Original',         width: 70 },
    { key: 'items',      header: 'Items / Summary',  width: 260 },
    { key: 'notes',      header: 'Extraction Notes', width: 260 },
    { key: 'fileId',     header: 'File ID',          width: 120, hidden: true },
    { key: 'originalId', header: 'Original File ID', width: 120, hidden: true }
  ];

  const work = {
    key: 'work',
    label: 'Work',
    idPrefix: 'W',
    folderName: 'Work',
    expensesSheet: 'Work Expenses',
    categoriesSheet: 'Work Categories',
    headerColor: '#1f3864',
    reimbursement: true,
    aiContext: 'These are work expenses paid personally by an employee; some will be submitted to the employer for reimbursement.',
    columns: [{ key: 'select', header: 'Select', width: 60 }]
      .concat(receiptCols)
      .concat([{ key: 'type', header: 'Type', width: 110 }])
      .concat(amountCols)
      .concat([
        { key: 'purpose',   header: 'Business Purpose', width: 220 },
        { key: 'site',      header: 'Site / Project',   width: 130 },
        { key: 'status',    header: 'Status',           width: 115 },
        { key: 'submitted', header: 'Submitted On',     width: 100 },
        { key: 'batch',     header: 'Batch ID',         width: 135 }
      ])
      .concat(fileCols),
    statuses: [STATUS.NEEDS_REVIEW, STATUS.UNSUBMITTED, STATUS.SUBMITTED,
               STATUS.REIMBURSED, STATUS.REJECTED, STATUS.TRACKED],
    // Category tab layout. "Default Type" decides whether a new receipt in
    // that category starts as Reimbursable or Track Only.
    categoryColumns: [
      { key: 'name',        header: 'Category',                          width: 190 },
      { key: 'type',        header: 'Default Type',                      width: 120 },
      { key: 'budget',      header: 'Monthly Budget',                    width: 120 },
      { key: 'description', header: 'Description (helps the AI choose)', width: 420 }
    ],
    defaultCategories: [
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
    ],
    /** Starting status for a cleanly-read receipt. */
    initialStatus: type => (type === TYPES.REIMBURSABLE ? STATUS.UNSUBMITTED : STATUS.TRACKED)
  };

  const home = {
    key: 'home',
    label: 'Home',
    idPrefix: 'H',
    folderName: 'Home',
    expensesSheet: 'Home Expenses',
    categoriesSheet: 'Home Categories',
    headerColor: '#274e13',
    reimbursement: false,
    aiContext: 'These are personal and household purchases.',
    columns: receiptCols
      .concat(amountCols)
      .concat([
        { key: 'memo',   header: 'Memo',   width: 220 },
        { key: 'status', header: 'Status', width: 115 }
      ])
      .concat(fileCols),
    statuses: [STATUS.NEEDS_REVIEW, STATUS.RECORDED],
    categoryColumns: [
      { key: 'name',        header: 'Category',                          width: 190 },
      { key: 'description', header: 'Description (helps the AI choose)', width: 420 }
    ],
    defaultCategories: [
      ['Groceries', 'Supermarkets and grocery stores, food for the home'],
      ['Dining Out', 'Restaurants, fast food, coffee shops, takeout, delivery'],
      ['Fuel', 'Gasoline/diesel purchases at a fuel station'],
      ['Auto', 'Car repairs, maintenance, parts, car wash, registration'],
      ['Household', 'Cleaning supplies, toiletries, paper goods, general household items'],
      ['Home Improvement', 'Hardware stores, tools, building materials, garden'],
      ['Utilities', 'Electric, water, gas, internet, phone bills'],
      ['Medical & Pharmacy', 'Doctor, dental, vision, pharmacy, prescriptions'],
      ['Clothing', 'Clothes, shoes, accessories'],
      ['Electronics', 'Computers, phones, gadgets, accessories for personal use'],
      ['Entertainment', 'Movies, events, games, streaming, hobbies'],
      ['Gifts & Donations', 'Gifts for others, charitable donations'],
      ['Other', 'Anything that does not fit another category']
    ],
    initialStatus: () => STATUS.RECORDED
  };

  BOOKS_CACHE_ = [work, home];
  return BOOKS_CACHE_;
}

/**
 * Returns one book by key ('work' or 'home').
 * @param {string} key
 * @return {Object}
 */
function getBook_(key) {
  const b = getBooks_().find(x => x.key === key);
  if (!b) throw new Error('Unknown book: ' + key);
  return b;
}
