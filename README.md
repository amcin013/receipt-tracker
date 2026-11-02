# Receipt Tracker

**Created:** 2026-10-30
**Author:** Allen McIntosh (with Claude)
**Stack:** Google Sheets + Apps Script (V8) + Google Drive + Gemini API + Gmail

You scan a receipt with your phone into a Drive folder. The script reads it with Gemini, renames and files the image, and adds a row to the sheet. Reimbursable receipts (fuel, parking, meals) can be batched into one email with the receipts attached. Track-only purchases (equipment, cables, software) feed a monthly budget summary.

---

## Files

| File | Purpose |
|---|---|
| `Config.gs` | Constants, column map, statuses, shared helpers |
| `Setup.gs` | Menu, first-time setup, folders, tabs, Summary, trigger |
| `Gemini.gs` | Sends the receipt to Gemini and parses the structured JSON reply |
| `Processor.gs` | Watches the Inbox folder, writes rows, renames and moves files |
| `Reimbursement.gs` | Builds the reimbursement email, mark-reimbursed, undo batch |
| `appsscript.json` | Manifest (time zone America/Chicago, V8) |

---

## One-time setup (about 10 minutes)

1. **Pick the account.** For work expenses, do all of this signed in to your **LHI Google account**. The receipts, sheet, and emails then all live in LHI's Workspace.
2. **Create the sheet.** Make a new Google Sheet named something like `Expense Tracker`.
3. **Add the code.** Go to **Extensions > Apps Script**. Delete the starter `Code.gs` and create five script files named `Config`, `Setup`, `Gemini`, `Processor`, and `Reimbursement`, pasting in the matching files.
   To edit the manifest, open **Project Settings (gear) > Show "appsscript.json" manifest file in editor**, then paste in `appsscript.json`. Save.
4. **Reload the sheet.** An **Expenses** menu appears.
5. **Run setup.** Go to **Expenses > Setup > Run first-time setup**, approve the permissions prompt (Drive, Sheets, Gmail, external requests), and setup will run.
   It creates:
   - Drive folders: `Expense Receipts/Inbox`, `Processed/YYYY/MM`, `Needs Attention`
   - Tabs: Expenses, Categories, Summary, Settings, Log
   - A trigger that checks the Inbox every 10 minutes
6. **Get a Gemini API key.** Go to aistudio.google.com, then **Get API key**. Paste it in via **Expenses > Setup > Set Gemini API key**. It's stored in Script Properties, not on a sheet.
7. **Fill in the Settings tab.** At minimum, set **Your Name** and **Reimbursement Email To**.
8. **Optional:** set monthly budgets in the Categories tab (column C), then run **Expenses > Setup > Rebuild Summary tab**.

### Privacy / compliance note

Receipts aren't PHI, so HIPAA doesn't really apply here. Still, they carry card last-4 digits and vendor history. Under Google's Gemini API terms, content sent through the **free tier** may be used to improve Google's products. The **paid tier** (billing enabled on the AI Studio / Cloud project) is not used that way. Check the current terms, and enable billing on the key if that matters for LHI. At a few receipts a week, the cost is negligible.

---

## Daily use

### Scanning (phone)

1. Open the **Google Drive** app, tap **+**, then **Scan**.
2. Capture the receipt. It auto-crops; add pages if it's long.
3. On save, choose the **Expense Receipts > Inbox** folder. Drive usually remembers the last folder you used.

Photos also work: share an image from your gallery to Drive and save it to Inbox. Online orders (Amazon, Newegg) don't need scanning. Download the invoice PDF and drop it in Inbox.

### What happens next

Within 10 minutes (or right away with **Expenses > Process Inbox now**):

- A row appears with date, vendor, category, subtotal, tax, total, card last-4, and gallons and price per gallon for fuel.
- The file is renamed, e.g. `2026-10-30_Shell_42.17_R-2026-0001.pdf`, and moved to `Processed/2026/10`.
- **Fuel, parking, meals, and lodging** start as **Reimbursable / Unsubmitted**.
- **Equipment, cables, software, and supplies** start as **Track Only / Tracked**.
- Anything blurry, missing a total or date, or that looks like a duplicate goes to **Needs Review**, and the file goes to `Needs Attention`.

### Fill in what a receipt can't tell the script

Add **Business Purpose** (e.g., "Travel to Round Rock site, PACS workstation install") and optionally **Site / Project**. If most fuel trips share the same purpose, set **Default Fuel Purpose** in Settings and it's filled in automatically.

### Submitting for reimbursement

1. Tick **Select** on the reimbursable rows.
2. Run **Expenses > Prepare reimbursement email (selected rows)**.
3. The script checks every row: it must be Reimbursable, have a total, a date, and a purpose (if required), and not already be submitted.
4. One email is built with a summary table, a grand total, and all receipts attached.
   - **Send Mode = Draft** (default): it lands in Gmail Drafts for you to review and send.
   - **Send Mode = Send**: it goes out immediately.
5. The rows are stamped **Submitted** with today's date and a **Batch ID**.
6. When the money arrives, tick those rows and run **Mark selected as Reimbursed**.

If you discard a draft, run **Undo a reimbursement batch...** with its Batch ID to reset those rows.

### Statuses

| Status | Meaning |
|---|---|
| Needs Review | Couldn't be read cleanly, or a possible duplicate. Check the receipt, fix the fields, then set Unsubmitted or Tracked |
| Unsubmitted | Reimbursable and not sent yet |
| Submitted | Included in a reimbursement email |
| Reimbursed | Paid back |
| Rejected | Denied (you can resubmit) |
| Tracked | Track-only expense; nothing to do |

### Budget tracking (Summary tab)

- **B1** picks the month. It follows the current month by default; type any date to view a different month.
- **Budget vs. actual** shows each category's budget, amount spent, amount remaining, and % used. Rows turn yellow at 85% and red when over.
- **Reimbursement snapshot** shows how much is unsubmitted and how much is submitted but not yet paid.
- **Monthly history** is one row per month with a column per category, so it's easy to chart.

---

## Customizing

- **Categories:** add, rename, or re-type rows on the Categories tab. The description column is sent to Gemini to help it choose. Run **Rebuild Summary tab** afterward.
- **Model:** change **Gemini Model** in Settings. The default `gemini-flash-latest` is Google's rolling alias for the current Flash model.
- **Columns:** the order is defined once in `COLUMNS` in `Config.gs`. Change it only **before** first setup, since existing rows won't be rearranged.

---

## Troubleshooting

| Symptom | Check |
|---|---|
| Nothing happens after scanning | Is the file in `Inbox`, not just "My Drive"? Check the **Log** tab, or run Process Inbox now |
| Log says "setup problem" | API key is wrong or missing, or the Gemini Model name is invalid. Files stay in Inbox and retry automatically once fixed |
| Rows show Needs Review often | Scan flatter and in better light; faded thermal receipts read poorly. Fix the fields by hand |
| Trigger stopped | Apps Script > Triggers. Re-run **Turn on auto-processing** |
| Email blocked: "attachments too large" | Gmail caps attachments at 25 MB. Submit in two batches |

---

## Possible next steps

- Import Amazon and other order confirmations automatically from Gmail (no scanning)
- Mileage log alongside fuel, if LHI reimburses per mile
- Charts on the Summary tab from the monthly history table

---

## Changelog

- **2026-10-30**: v1.0. Initial build: Drive inbox processing, Gemini extraction, budget summary, reimbursement email batching.
