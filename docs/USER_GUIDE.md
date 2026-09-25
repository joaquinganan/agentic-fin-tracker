# User guide

Setting up and using the tracker in your own Google Sheet.

## 🚀 QUICK START

### Step 1: Create the Google Sheet first

1. Create a new Google Sheet
2. Name it "Financial Tracker"

### Step 2: Open Apps Script FROM the Sheet (important)

This script uses `SpreadsheetApp.getActiveSpreadsheet()` and an `onOpen()` menu trigger, both of which only work when the script is **bound to the Sheet** — not a standalone project. There is no Spreadsheet ID to paste anywhere; do NOT create the project at script.google.com directly.

1. In your new Google Sheet, go to the menu: **Extensions > Apps Script**
   - This opens an Apps Script project that's automatically linked to this Sheet.
2. Delete the default `Code.gs` content
3. Create 6 script files (use the **+** next to "Files") named exactly like the files in `src/`
   (or push them with [clasp](https://github.com/google/clasp) — see the main README):
   - `01_main.gs`
   - `02_categorizer.gs`
   - `03_gmailMonitor.gs`
   - `04_sheetsWriter.gs`
   - `05_dailySummary.gs` (v1.1.24)
   - `06_monthlySummary.gs` (v1.1.26)
4. Paste the matching content into each file
5. Save the project (Ctrl+S / Cmd+S)

### Step 3: Run Setup Wizard

1. In Apps Script, go to `01_main.gs`
2. Find `function onOpen()` and click the play button (▶)
3. **Authorize** when prompted (grant permissions)
4. Refresh your Google Sheet
5. Click menu: **📊 Tracker** > **🔧 Setup Wizard**
6. Fill in your information:
   - Email, income currency, monthly gross salary and any other monthly income (no deductions)
   - Deductions: automatic (DOP salary, DR payroll rules) or entered manually
   - Banks to track (LAFISE, BANESCO, BHD, POPULAR, BDI) and, for each, your credit card with its statement and payment days
   - Optionally, summary emails: daily and/or monthly (on the 1st), recipient, hour, and what the daily one includes
7. Click **✅ Save Configuration**

### Step 4: Test Gmail Monitor

1. In Google Sheet, click: **📊 Tracker** > **🔄 Monitor Gmail Now**
2. System searches your banks' emails for the current month (read or unread)
3. Parses and imports them automatically
4. Check the **Transactions** sheet for results

---

## 📋 WHAT IT DOES

Reads your banks' notification emails (LAFISE, BANESCO, BHD, POPULAR; BDI via a generic
fallback until a real sample exists), extracts each transaction, categorizes it, and keeps
a Google Sheet with a live Dashboard. Runs every day at 6:00 AM and on demand.

```
Gmail (by sender) → parse each message → filters (promotional / non-transactional / declined)
  → Type (subject first, then body) → bank+type extractor (generic fallback if none)
  → Custom Rules, then default categories → batch save (dedup by Gmail message id)
  → sort → recategorize + rebuild Raw_<BANK> / Bank Transfers → Dashboard
```

**Transactions is the single source of truth.** `Raw_<BANK>` and `Bank Transfers` are
rebuilt from it on every run (your **Notes** in `Raw_<BANK>` column G are preserved).

---

## 📊 SHEETS

| Sheet | What it holds |
|---|---|
| Dashboard | Month by name (**Current month** follows today) and year, exchange rates with the date they were last edited, KPI cards (net income · spent · remaining · % of income · transfers to review), categories for the month with in-cell bars, income & deductions, fixed vs. variable, spend by bank, month-by-month table + chart, the year as a heat map, credit cards, and "Which card for what" (the LAFISE, BANESCO and BHD cashback programs side by side). Yellow cells are yours; they live in named ranges (`DASH_MONTH`, `DASH_YEAR`, `RATE_USD/EUR/COP`, `RATE_UPDATED`, `DASH_CARDS`) and survive every rebuild. |
| Transactions | `Date · Bank · Merchant · Amount · Currency · Category · Description · Email Subject · Timestamp (hidden) · IsCredit · IsCashback (hidden) · Type · MessageId (hidden) · TxRef (hidden — the bank's own transaction id)` |
| Bank Transfers | Every Type = Transfer row |
| Raw_<BANK> | Every Type = Transaction row for that bank, plus your Notes |
| Custom Rules | `UserEmail · Category · Keyword` — checked before the defaults, for any Type. Category **Exclude** leaves a row out of every total (self-transfers). A category that isn't a default one gets its own Dashboard row. |
| Configuration | Key/value. The Dashboard reads it through named ranges (`CFG_MONTHLY_INCOME`, `CFG_ARS`, `CFG_AFP`, `CFG_TAX_RATE`, `CFG_INCOME_CURRENCY`), so row order no longer matters. |
| Categories | Auto-generated reference: default keywords + your Custom Rules |

**Type:** `Transaction` · `Transfer` · `Card Payment` · `Cashback`. Only Transaction (and any
row a Custom Rule categorizes) counts as spending; Card Payment and Cashback never do.
**IsCredit** = money coming in (Card Payment, Cashback, or an explicit crédito/depósito in
that transaction's own text).

---

## 📌 CATEGORIES

11 defaults, evaluated **in this order** (first match wins; order matters —
e.g. Dining is checked before Vehicle Gas so "BONJOUR TOTAL" isn't gas):
Rent · Gym + Calisthenics · Telecommunications · Streaming & Subscriptions · Electricity ·
Groceries + Barbershop · Dining/Delivery + Entertainment + Other (also the fallback) ·
Vehicle Gas · Health + Vet + Pharmacy · Education · Transportation.
Fixed: Rent, Gym, Telecommunications, Streaming. Full keyword lists: the **Categories** sheet.

**Keyword matching rule (v1.1.19):** keywords of 4 letters/digits or fewer (`BAR`, `BUS`,
`UBER`, `ACH`…) match whole words only; longer ones match anywhere in the text
(`SUPER` → SUPERMERCADO). The same rule is used by categories, Custom Rules, Type detection
and the email filters.

---

## 🔍 MENU (📊 Tracker)

- **🔧 Setup Wizard** — opens pre-filled with your current values; validates before saving.
- **🔄 Monitor Gmail Now** — same as the 6 AM run: from the 1st of *yesterday's* month
  through today, both included.
- **📅 Monitor by Date Range** — any inclusive range; the window closes itself and the
  summary appears when the run ends. Use it for backlogs (a quarter at a time).
- **📊 Dashboard** — rebuilds the layout (numbers are live formulas either way).
- **🔁 Recategorize Saved Transactions** — re-applies the current rules to every row
  (runs automatically at the end of every monitor run). Never changes merchant names.
- **⚙️ View Config** · **🗑️ Reset System**

Only one run can happen at a time; a second one is refused with a message.

---

## 🔧 TECHNICAL NOTES

- **Search:** `(from:<bank senders>) after:<epoch> before:<epoch>` — exact bounds in the
  script's time zone, end date included, paginated up to 2,000 threads per run. Messages
  outside the range inside a matched thread are skipped.
- **Duplicates:** Gmail message id + item index (`<id>_0`, `<id>_1`…). Rows saved before
  v1.1.13 (no MessageId) are matched by date + bank + amount.
- **Gmail side effects:** processed threads are marked read and labeled **Procesado** (the
  label is created on first use). A thread with an email that couldn't be parsed is left
  **unread** so you notice it.
- **Time zone:** `appsscript.json` → `"timeZone": "America/Santo_Domingo"` must match the
  spreadsheet's (File → Settings).

---

## 🧪 TESTS (v1.1.19)

`tests/` runs the real `.gs` files in Node, in one shared scope like Apps Script, against
anonymized copies of real bank emails plus a strict Sheets/Gmail mock (unknown API methods
and out-of-bounds writes throw). Requires Node 18+:

```
node --test tests/*.test.js
```

Every new real email sample should become a fixture in `tests/fixtures/` (see
`tests/README.md`). `GS_ROOT=/path/to/other/copy node --test tests/*.test.js` runs the same
suite against another copy of the code.

---

## 🚨 TROUBLESHOOTING

- **Check the version first:** every run logs `SCRIPT_VERSION` (View → Executions). If it's
  not the latest, the fix you're testing isn't deployed.
- **A row still shows an old problem after an update:** saved rows aren't re-parsed. Delete
  the row, then run **📅 Monitor by Date Range** for that date (a re-run without deleting is
  skipped as a duplicate). Type and Category, however, are refreshed by Recategorize.
- **"Another run is already in progress":** wait for the running one's summary. A run killed
  by the 6-minute limit releases the lock automatically.
- **"Could not parse" in the summary:** the thread was left unread in Gmail. Open the
  execution log, find `Amount not found | …`, and share that email's .eml.

---

## 📝 OPEN ITEMS

- A real **bundled multi-transaction BHD transfer** .eml (the "no Beneficiario" fallback is still inferred).
- A real **LAFISE transfer** .eml (current extractor is a stopgap).
- A real **BDI** consumo sample (BDI uses the generic fallback).
- A real **cashback** email (IsCashback has never been YES on real data).
- A real **multi-row BHD email with a declined row** (M1 is covered by a synthetic fixture only).
- Exchange rates are manual; EUR/COP defaults are approximations.
- A shared setup for a friend's banks (multi-user).

---
