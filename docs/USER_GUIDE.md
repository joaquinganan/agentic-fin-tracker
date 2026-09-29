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
2. Open **[dist/FinancialTracker.gs](../dist/FinancialTracker.gs)** in the repository, copy all of it, and paste it
   into `Code.gs`, replacing what's there. That one file is the whole tracker (you can rename it `FinancialTracker`).
3. Save the project (Ctrl+S / Cmd+S) and reload the Google Sheet: the **📊 Tracker** menu appears.

**Updating:** open `FinancialTracker.gs`, select everything (Ctrl+A), paste the new version, save. The first line says
which version it is.

**Coming from the seven files (`01_main.gs` … `07_investments.gs`)?** Delete all seven (⋮ next to each file › Delete),
then add `FinancialTracker.gs` as above. Don't keep both: the same code twice stops the project from loading (the
📊 Tracker menu disappears). Your data, settings and scheduled runs stay as they are.

### 📘 Start here — a checklist that checks itself

**📊 Tracker › 📘 Start here** opens a panel beside the sheet with every step: settings, the daily update, reading this
year's emails, emails waiting in Unrecognized, transfers without a category and, if you track investments, your
positions, where each account's return starts and balances to update. Every step is checked against the sheet — nothing
to tick by hand — so it is also the health check: when something breaks later, it shows up there again with a button
for what to do. Open it after the Setup Wizard, and whenever something looks off.

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

## Investments

1. Add the file `07_investments.gs` to the Apps Script project (like the others).
2. **Investment Accounts** (created on the first run): one row per account. *Deposit keyword* identifies the bank
   transfers that fund it (for HAPI it is pre-filled with its DR collection account).
3. **Snapshot** each broker account once: **📊 Tracker › 📋 Paste Broker Positions**. In the broker's web app, open the
   portfolio, select from *Total balance* down to the last asset, copy and paste; check the preview and save. (It can
   also be typed in the Investment Ledger: one `Snapshot` row per ticker — Quantity, Price, Amount = cost basis — plus a
   `CASH` row.) Positions start from it and add only the movements after that day.
4. Funds and pensions: **📊 Tracker › ➕ Add Balance or Deposit** whenever you get a statement (units × unit
   price, or the balance). It adds a new `Valuation` row each time — Holdings shows the latest of each account.
   Saving the same account and date again updates that balance. To measure a return, add an earlier statement too
   (e.g. the balance on January 1): the earliest balance is where the account's history starts.
5. **📊 Tracker › 📈 Refresh Investments** rebuilds Holdings. Regular runs read broker emails and bank deposits on
   their own. Holdings is rebuilt every time — don't edit it by hand; edit the ledger.
6. Every run records the day's values in **Portfolio History**. Holdings shows the return of each account since it
   started being tracked, and a chart of the total; the daily and monthly emails include an Investments section
   (turn it off in the Setup Wizard). Returns are annualized once an account has 180 days of history.
7. A new account (e.g. IBKR) needs no Snapshot if it starts empty: its movements are added from zero. Add it to
   Investment Accounts with the keyword of the transfers that fund it.

## Unrecognized emails

The **Unrecognized** sheet lists every bank or broker email the tracker couldn't read, with the reason, what the email
says and a link to open it in Gmail. Send those emails (Gmail › ⋮ › Show original › Download) to get their format
supported; set Status to *Ignore* for ones that don't matter. Rows disappear by themselves once a later version reads
the email.

## Long runs

Apps Script stops any run at 6 minutes. The tracker stops reading emails before that, saves what it read and tells you
in the summary ("⏸ Stopped reading early…", "⏭ Left for the next run…"): run the same range again until those lines are
gone — emails already read are skipped, so each run gets further.

## Categories you set by hand

You can set or change a category directly in **Transactions** or **Bank Transfers**: the tracker keeps it — recategorizing
and the daily run no longer overwrite it. To go back to the automatic category, clear the cell. For something that should
apply to every future transaction too, add a Custom Rule instead. (Edits in the Raw_ sheets aren't kept: those are copies.)

## Deposits and the return

An account's gain is its value now − its value at the start − the money put in since. Deposits are recorded by
themselves when they're bank transfers matching the account's keyword (Investment Accounts); anything else — another
route, a bank whose transfers aren't read yet, pension contributions — add with **➕ Add Balance or Deposit › A deposit**.
If purchases since the start exceed the recorded deposits, Holdings › Performance warns you: a missing deposit counts as gain.
HAPI emails a notice for every completed deposit, without the amount: the tracker keeps them and lists in
**Unrecognized** each one with no deposit recorded near its date — add that deposit's amount and the row goes away.

## Incoming transfers: money paid back to you

Money you receive lands in **Incoming Transfers**, without a category (highlighted). Give it the category it pays back
— a roommate's share of the rent → *Rent* — and it is taken off what you spent there, on the Dashboard and in the
summaries; *Exclude* for money that pays nothing back. When the transfer names its sender (LAFISE does), a Custom Rule
on that name categorizes it by itself. Transfers from your own account are recognized and excluded.

Banesco doesn't notify most incoming transfers, but its monthly savings statement (a PDF) lists them — as "Ach Ibanking",
without the sender. The tracker reads the PDF itself (nothing to turn on). What's read must add up to the statement's own
totals, or nothing is taken and it shows in Unrecognized.

A category typed in Incoming Transfers or Bank Transfers applies at once (the Dashboard doesn't wait for the next
update); each transfer counts in the month of its date. A transfer the tracker didn't see can be typed in an empty row
of Incoming Transfers — date (yyyy-mm-dd), bank (optional), from, category, amount (positive; DOP if no currency) — and
is saved at the next update or 🔁 Recategorize; missing a date or an amount, it's listed in Unrecognized instead.
