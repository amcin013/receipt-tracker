# Receipt Tracker

**Created:** 2026-10-30
**Updated:** 2026-11-02 (v2.0)
**Author:** Allen McIntosh (with Claude)
**Stack:** Google Sheets + Apps Script (V8) + Google Drive + Gemini API + Gmail

You scan a receipt with your phone into a Drive folder: **Work Inbox** or **Home Inbox**. The script reads it with Gemini, keeps an untouched copy of the original, renames and files the working copy, and adds a row to that book's tab. Work and Home are fully separate "books" with their own folders, tabs, categories, and receipt numbers. Nothing is guessed: the inbox you scan into decides the book.

- **Work** keeps the full v1 feature set. Reimbursable receipts (fuel, parking, meals) can be batched into one email with the receipts attached. Track-only purchases (equipment, cables, software) feed a monthly budget summary.
- **Home** is storage and data only. Every household receipt is filed and recorded in a clean table, ready for a home budget sheet to pull from later.

---

## Files

All code lives in `src/`.

| File | Purpose |
|---|---|
| `Config.gs` | Generic constants, statuses, column lookup, shared helpers |
| `Books.gs` | Work and Home definitions: columns, categories, statuses, folder names, ID prefixes |
| `Setup.gs` | Menu, first-time setup, folders, tabs, Work Summary, trigger |
| `Gemini.gs` | Sends the receipt to Gemini and parses the structured JSON reply |
| `Processor.gs` | Watches both Inbox folders, archives originals, writes rows, renames and moves files |
| `Reimbursement.gs` | Work only: builds the reimbursement email, mark-reimbursed, undo batch |
| `appsscript.json` | Manifest (time zone America/Chicago, V8) |

---

## Where everything lives

### Account

Everything runs under your **personal Google account**, not LHI's Workspace. That way:

- You own every receipt, work and home, even if you change jobs. This is your "personal copy of work receipts."
- LHI admins can't see your household purchases.
- LHI still gets what it needs: each reimbursement email carries the receipts as attachments.

Reimbursement emails go from your personal Gmail to your LHI inbox, and you forward them from there (see "How reimbursement emails reach LHI" below).

### Drive layout (created by setup)

```
My Drive/
  Receipts/
    Work/
      Inbox/                 <- scan work receipts here
      Processed/2026/10/     <- renamed working copies (linked as "Receipt")
      Needs Attention/       <- unreadable, low-confidence, or possible duplicates
      Originals/2026/10/     <- untouched copy of every scan (linked as "Original")
    Home/
      Inbox/                 <- scan home receipts here
      Processed/...
      Needs Attention/
      Originals/...
```

### Spreadsheet tabs

| Tab | What it holds |
|---|---|
| Work Expenses | One row per work receipt: reimbursement status, purpose, batch |
| Home Expenses | One row per home receipt, with a Memo column for your own notes |
| Work Summary | Work budget vs. actual by month, reimbursement totals, monthly history |
| Work Categories / Home Categories | Each book's categories. Descriptions help Gemini choose |
| Settings | Email, send mode, model, and other options |
| Log | Processing history for both books |

**Receipt IDs** are separate per book and reset each year: `W-2026-0001` and `H-2026-0001`. A Work row and a Home row never share a number.

---

## One-time setup (about 10 minutes)

1. **Use your personal Google account.** If you already set up v1 in your LHI account, open that sheet and run **Expenses > Setup > Turn off auto-processing** there first, so two copies aren't running.
2. **Create the sheet.** Make a new Google Sheet named something like `Receipts`.
3. **Add the code.** Go to **Extensions > Apps Script**. Delete the starter `Code.gs` and create six script files named `Config`, `Books`, `Setup`, `Gemini`, `Processor`, and `Reimbursement`, pasting in the matching files from `src/`. File order doesn't matter.
   To edit the manifest, open **Project Settings (gear) > Show "appsscript.json" manifest file in editor**, then paste in `src/appsscript.json`. Save.
4. **Reload the sheet.** An **Expenses** menu appears.
5. **Run setup.** Go to **Expenses > Setup > Run first-time setup** and approve the permissions prompt (Drive, Sheets, Gmail, external requests). After you approve, setup keeps running in the background. Wait for the **Setup complete** popup (up to a minute) before clicking anything else. Setup creates the folders and tabs above and a trigger that checks both Inboxes every 10 minutes.
6. **Get a Gemini API key.** Go to aistudio.google.com, then **Get API key**. Paste it in via **Expenses > Setup > Set Gemini API key**. It's stored in Script Properties, not on a sheet.
7. **Fill in the Settings tab.** At minimum, set **Your Name** and **Reimbursement Email To**.
8. **Bookmark both Inboxes on your phone.** **Expenses > Open Inbox folders** shows both links.
9. **Optional:** set monthly Work budgets in Work Categories (column C), then run **Expenses > Setup > Rebuild Work Summary tab**.

### Settings reference

| Key | Default | Notes |
|---|---|---|
| Your Name | | Signs reimbursement emails |
| Reimbursement Email To | | Required for Work reimbursement. Usually your own LHI address |
| Reimbursement Email CC | | Optional, e.g. your LHI address |
| Send From Alias | | Optional, normally blank. See "How reimbursement emails reach LHI" |
| Subject Prefix | Expense Reimbursement Request | |
| Send Mode | Draft | Draft = review in Gmail first; Send = send immediately |
| Require Business Purpose | Yes | Blocks Work submission if Purpose is blank |
| Default Fuel Purpose | | Auto-fills Purpose on new Work fuel receipts |
| Keep Original Copies | Yes | No = skip the Originals archive |
| Gemini Model | gemini-flash-latest | |

### How reimbursement emails reach LHI

The recommended setup is to send to yourself and forward:

| Setting | Value |
|---|---|
| Reimbursement Email To | your LHI address |
| Send From Alias | blank |
| Send Mode | Send |

The script emails the batch from your personal Gmail to your LHI inbox. From there you forward it, or copy it into a new email and edit it, to whoever approves expenses. Rows are stamped **Submitted** when the script sends it to your LHI inbox, so if you decide not to forward a batch, run **Work: Undo a reimbursement batch** with the Batch ID at the bottom of the email.

**Optional: sending directly as your LHI address.** You can add the LHI address as a Gmail "Send mail as" alias (personal Gmail > Settings > Accounts and Import > Send mail as). If Gmail asks for SMTP details, use smtp.gmail.com, port 587, and an **app password** created while signed in to the LHI account; your regular password is rejected. Workspace admins often turn app passwords off. Once the alias is verified, put the address in **Send From Alias**. The script refuses to send if the alias isn't verified.

### Privacy note

Receipts aren't PHI under HIPAA, since HIPAA covers providers, plans, and their business associates, not your own records. Home receipts can still be sensitive: pharmacy and medical receipts may list drug or service names, and everything carries card last-4 digits and shopping history. Two things to know:

- **Gemini terms.** Content sent through the Gemini API **free tier** may be used to improve Google's products. The **paid tier** (billing enabled on the AI Studio / Cloud project) is not. With Home receipts going through it too, enabling billing is worth it. At a few receipts a week, the cost is negligible. Check the current terms.
- **Don't share the workbook.** Work and Home are separate tabs in one file, and Sheets sharing is all-or-nothing per file. If you ever need to give someone at LHI access, send them a reimbursement email or export, not the spreadsheet.

Also check whether LHI policy requires work receipts to be kept in company systems. The emails cover most policies, but it's worth confirming.

---

## Daily use

### Scanning (phone)

1. Open the **Google Drive** app, tap **+**, then **Scan**.
2. Capture the receipt. It auto-crops; add pages if it's long.
3. On save, choose **Receipts > Work > Inbox** or **Receipts > Home > Inbox**. Drive remembers the last folder, so **check it each time you switch books**.

Photos also work: share an image from your gallery to Drive and save it to an Inbox. Online orders (Amazon, Newegg) don't need scanning. Download the invoice PDF and drop it in the right Inbox.

**Scanned into the wrong Inbox?** Delete the row first, move the file from that book's `Processed` back into the correct Inbox, and it'll be recorded again on the next run with a new ID. You can delete the stray `Originals` copy too. If you scan the same receipt into both inboxes, the second one is flagged **Possible duplicate of W-...** (or H-...), so cross-book mistakes are easy to spot.

### What happens next

Within 10 minutes (or right away with **Expenses > Process both Inboxes now**):

- An untouched copy of the scan is saved to `<Book>/Originals/YYYY/MM` as `W-2026-0001_<original name>`.
- A row appears on the book's tab with date, vendor, category, subtotal, tax, total, card last-4, gallons and price per gallon for fuel, and **Receipt** and **Original** links.
- The working file is renamed, e.g. `2026-10-30_Shell_42.17_W-2026-0001.pdf`, and moved to `<Book>/Processed/YYYY/MM`.
- **Work:** fuel, parking, meals, and lodging start as **Reimbursable / Unsubmitted**. Equipment, cables, software, and supplies start as **Track Only / Tracked**.
- **Home:** clean reads start as **Recorded**.
- In either book, anything blurry, missing a total or date, or that looks like a duplicate goes to **Needs Review**, and the file goes to that book's `Needs Attention`.

### Work: fill in what a receipt can't tell the script

Add **Business Purpose** (e.g., "Travel to Round Rock site, PACS workstation install") and optionally **Site / Project**. If most fuel trips share the same purpose, set **Default Fuel Purpose** in Settings and it's filled in automatically.

### Work: submitting for reimbursement

1. On **Work Expenses**, tick **Select** on the reimbursable rows.
2. Run **Expenses > Work: Prepare reimbursement email (selected rows)**.
3. The script checks every row: it must be Reimbursable, have a total, a date, and a purpose (if required), and not already be submitted.
4. One email is built with a summary table, a grand total, and all receipts attached.
   - **Send Mode = Draft** (default): it lands in Gmail Drafts for you to review and send.
   - **Send Mode = Send**: it goes out immediately.
5. The rows are stamped **Submitted** with today's date and a **Batch ID**.
6. When the money arrives, tick those rows and run **Work: Mark selected as Reimbursed**.

If you discard a draft, run **Work: Undo a reimbursement batch...** with its Batch ID to reset those rows.

### Statuses

| Status | Book | Meaning |
|---|---|---|
| Needs Review | Both | Couldn't be read cleanly, or a possible duplicate. Check the receipt, fix the fields, then set the normal status |
| Unsubmitted | Work | Reimbursable and not sent yet |
| Submitted | Work | Included in a reimbursement email |
| Reimbursed | Work | Paid back |
| Rejected | Work | Denied (you can resubmit) |
| Tracked | Work | Track-only expense; nothing to do |
| Recorded | Home | Stored and read cleanly |

### Work budget tracking (Work Summary tab)

- **B1** picks the month. It follows the current month by default; type any date to view a different month.
- **Budget vs. actual** shows each Work category's budget, amount spent, amount remaining, and % used. Rows turn yellow at 85% and red when over.
- **Reimbursement snapshot** shows how much is unsubmitted and how much is submitted but not yet paid.
- **Monthly history** is one row per month with a column per category, so it's easy to chart.

---

## Using Home data in your budget sheet

The Home Expenses tab is a plain table (one row per receipt, fixed headers), so getting data out is easy:

- **Manual, for now:** filter Home Expenses by month and copy the Date, Vendor, Category, and Total columns.
- **Later, live link:** in your budget spreadsheet, `=IMPORTRANGE("<this sheet's URL>", "'Home Expenses'!A:H")` pulls the data in. Wrap it in `QUERY` or `SUMIFS` to total by month and category. If your budget uses different category names, rename the rows on Home Categories to match before receipts pile up.

---

## Customizing

- **Categories:** add, rename, or re-type rows on Work Categories or Home Categories. The description column is sent to Gemini to help it choose. After changing Work categories, run **Rebuild Work Summary tab**.
- **Model:** change **Gemini Model** in Settings. The default `gemini-flash-latest` is Google's rolling alias for the current Flash model.
- **Columns:** each book's column order is defined in `Books.gs`. Change it only **before** first setup, since existing rows won't be rearranged.

---

## Troubleshooting

| Symptom | Check |
|---|---|
| Nothing happens after scanning | Is the file in `Receipts/Work/Inbox` or `Receipts/Home/Inbox`, not just "My Drive" or the book folder? Check the **Log** tab, or run Process both Inboxes now |
| Receipt landed in the wrong book | See "Scanned into the wrong Inbox?" above |
| File showed up in Needs Attention with no new row | The Log says "Already recorded as ... in the other book". You moved it between Inboxes without deleting the old row. Delete that row, then drop the file in the right Inbox again |
| Log says "setup problem" | API key is wrong or missing, or the Gemini Model name is invalid. Files stay in their Inbox and retry automatically once fixed |
| Rows show Needs Review often | Scan flatter and in better light; faded thermal receipts read poorly. Fix the fields by hand |
| "Send From Alias not set up" | Add the address under Gmail > Settings > Accounts > "Send mail as", or clear the setting |
| Trigger stopped | Apps Script > Triggers. Re-run **Turn on auto-processing** |
| Email blocked: "attachments too large" | Gmail caps attachments at 25 MB. Submit in two batches |

---

## Development notes

- **Version control:** `src/` mirrors the Apps Script project. To push from this repo instead of copy-pasting, use [clasp](https://github.com/google/clasp): `npm i -g @google/clasp`, `clasp login`, then `clasp clone <scriptId> --rootDir src` once, and `clasp push` after each change. `.clasp.json` is git-ignored because it holds your script ID.
- **Load order:** book and settings definitions are built inside functions (`getBooks_()`, `getDefaultSettings_()`) rather than as top-level constants, so the code works no matter what order Apps Script or clasp loads the files.
- **Crash safety:** each receipt's row is written before its file is moved. If a move fails, the next run finds the row by File ID and finishes the move, so a receipt is never filed without a row.

---

## Possible next steps

- Connect Home Expenses to the home budget spreadsheet (IMPORTRANGE or a shared category list)
- Import Amazon and other order confirmations automatically from Gmail (no scanning)
- Mileage log alongside fuel, if LHI reimburses per mile
- Charts on the Work Summary tab from the monthly history table

---

## Changelog

- **2026-11-02**: v2.0.2. Setup holds a lock so overlapping runs can't duplicate anything, and it removes duplicate Settings rows left by earlier runs. When a key appears twice, the first row now wins.
- **2026-11-02**: v2.0.1 (docs). Default workflow is now: email the batch to your own LHI inbox, then forward it. The alias is optional.
- **2026-11-02**: v2.0. Separate Work and Home books (own Inbox, folders, tabs, categories, and ID series), with the Inbox deciding the book. Untouched originals archived per book. Runs from your personal Google account. Duplicate check spans both books. Optional Send From Alias for reimbursement emails. Each row is now written before its file is moved; a file seen again is filed by its existing row's book and status, and a file already recorded in the other book is surfaced in Needs Attention instead of filed silently. Extraction notes are guarded against formula injection.
- **2026-10-30**: v1.0. Initial build: Drive inbox processing, Gemini extraction, budget summary, reimbursement email batching.
