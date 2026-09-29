# Changelog

Every release, newest first. Each entry says what was wrong, how it was found and what changed.

### v1.1.53 (Sept 29, 2026) — Incoming Transfers treated like Bank Transfers everywhere
- **Reported:** Incoming Transfers wasn't formatted with the rest after recategorizing. Its styling was applied (it uses
  the sheet layouts), but three places list sheets by name and v1.1.51 added it to none: fitting the columns (so it
  looked unformatted — narrow, cut-off columns), sorting by date, and 🗑️ Reset System (a derived sheet left behind with
  old data). A test now requires it to get what Bank Transfers gets — same conditional formats, same widths for the same
  content, the same order — and to be deleted by the reset. Tests: 179, on `src/` and on the single file.

### v1.1.52 (Sept 29, 2026) — statements read by the tracker itself; LAFISE announcements filtered
- **Reported:** every Banesco statement ended in Unrecognized ("totals were not found · 0 row(s) read"). v1.1.51 had
  Google Drive convert the PDF to a Doc — the one step that couldn't be tested outside Apps Script — and on the real
  statements the text it returned had neither totals nor rows. The PDF is now read by the tracker (new `09_pdfText.gs`):
  it inflates the page content (a pure-JS FlateDecode), follows the text operators with their positions (cm, Tm, Td,
  TJ spacing, Tc/Tw/Tz), decodes WinAnsi, /Differences glyph names and Type0 ToUnicode maps, and rebuilds the lines
  with each glyph's real width. On the real statement: 20 rows, totals matching, the four incoming transfers — through a
  whole run, in memory. No Drive API service, no extra permissions.
- A first version joined pieces by an estimated width and split words ("B alance", "C réd itos"); real glyph widths
  fixed it. The synthetic PDF fixture (`tests/fixtures/statements/make_synthetic_pdf.py`, invented data) is written the
  way the real file is — pieces placed at real widths, words spaced by TJ numbers with no space character, a Type0
  title — so a reader ignoring either fails (mutation-checked). Inflate is checked against Node's zlib (stored, fixed
  and dynamic blocks; wrapped and raw; up to 1 MB).
- **Reported:** LAFISE's "HORARIO TRANSFERENCIAS PAGOS AL INSTANTE" announcement (from the incoming-transfers sender)
  landed in Unrecognized as "Amount not found"; announcements are now filtered. Tests: 177, on `src/` and on the single
  file.

### v1.1.51 (Sept 28, 2026) — incoming transfers: money paid back reduces what you spent
- **Requested:** shared costs (rent, utilities) are paid with the user's cards and a roommate transfers their share —
  the Dashboard should show what was really spent. Money received is saved as a **negative** row of the new Type
  **Incoming**, the way reversals are: given a category it reduces that category everywhere (Dashboard formulas already
  net by category; the daily and monthly summaries now count Incoming rows). Without one it counts nowhere.
- Nothing is decided for the user (no rules for anyone's own situation): categories come from Custom Rules on the
  sender's name, or are typed in the new **Incoming Transfers** sheet (same layout as Bank Transfers, edits synced back
  to Transactions and kept by recategorizing). The one automatic rule is generic: money from your own account (the
  sender is the account holder — first two names, accents and cut names aside) is Exclude.
- **LAFISE** "TRANSFERENCIA ENTRANTE" emails come from a third address, now searched. Their body says "TRANSFERENCIA …
  RECIBIDA" and "PAGOS AL INSTANTE" (Transfer keywords): the subject now types them first, or they'd read as sent.
- **Banesco** doesn't notify most incoming transfers; its monthly savings statement (PDF) lists them. With the Drive API
  service on, the PDF is converted to a temporary Google Doc, read and trashed. Rows are read from the text in any layout
  (columns, one line, one cell per line); credit or debit comes from the running balance; and everything must add up to
  the statement's own totals (credits, debits, closing balance) or nothing is taken and it goes to Unrecognized. Credits
  by ACH ("Ach Ibanking", no sender) and LBTR ("Lbtr <name>") are the incoming transfers; interest is not. Checked on a
  real statement (20 rows, totals matching); the fixtures are synthetic.
- 📘 Start here: incoming transfers without a category; the Drive API when a tracked bank has a statement to read.
- Tests: 172, on `src/` and on the single file (mutation-checked: incoming typed as sent, the credit total check, own
  money after recategorizing, the summary netting).

### v1.1.50 (Sept 28, 2026) — Holdings named a menu item that no longer existed
- Holdings' note and its empty "Other accounts" row still sent people to "📊 Tracker › ➕ Add Fund / Pension Balance",
  renamed "➕ Add Balance or Deposit" in v1.1.40. Found reviewing a frame of the new demo video. A test now checks every
  "📊 Tracker › …" the tracker shows against the real menu (it fails with the old name back). Tests: 166.

### v1.1.49 (Sept 28, 2026) — a Spanish guide; deleted rows come back; Reset System starts over for real
- **docs/GUIA.md**, the user guide in Spanish (install, setup, 📘 Start here, daily use, categories, investments with
  HAPI, updating, common problems, adding a bank, privacy), linked from the README and from 📘 Start here. A test checks
  that it names every menu item exactly as the menu shows it and that its links resolve — so it can't fall behind.
- Found while documenting: the read log (v1.1.43) also recorded emails that WERE saved, so a transaction row you deleted
  was never read again — and after 🗑️ Reset System, re-reading the year skipped every email. The log now holds only
  emails with nothing to save (statements, promotions, declined); saved ones are recognized by Transactions itself.
- **🗑️ Reset System** also deletes Unrecognized, the read log and the "last update" record, and its confirmation says
  what's kept (the investment tabs: what you typed, which no email can bring back) and that it can't be undone. The
  guide's first draft had described it as only clearing settings — checked against the code before publishing.
- **Reported:** no investments in the monthly report — "Send Monthly Summary Now" reports the previous month, which was
  before the investment history began, so the section was left out without a word. It now says when the history starts
  and which summary will be the first with figures.
- Tests: 165, on `src/` and on the single file (mutation-checked: the read log, the reset, the guide).

### v1.1.48 (Sept 28, 2026) — paste the broker's portfolio screen to make a Snapshot
- **📊 Tracker › 📋 Paste Broker Positions** (also from 📘 Start here): paste the broker's portfolio screen as copied from
  the web app — plain, or with links — and the tracker reads each asset's ticker, quantity, value and gain into a
  Snapshot (cost = value − gain, price = value ÷ quantity, tiny prices kept exact) plus the cash from "Total money".
  A preview comes first; Save is enabled only after it, and again only after previewing any edit.
- "Total assets" in the paste is checked against the positions' sum: a position left out of the copy is named with the
  amount missing (the way a missing SHIB-sized position had to be found by hand before), and Save becomes "Save anyway".
- Saving the same account and day again replaces that snapshot instead of adding a second one. Tests: 161, on `src/`
  and on the single file (mutation-checked: the total check, same-day replacement); the dialog was also exercised in a
  real browser.

### v1.1.47 (Sept 28, 2026) — 📘 Start here: a checklist that checks itself
- New **📊 Tracker › 📘 Start here** (first in the menu): a sidebar with every step of setting up and keeping the tracker
  healthy — settings, the daily update and summary emails being scheduled, this year's emails read (and whether the last
  run stopped early, failed or is over a day old), emails waiting in Unrecognized, transfers this month without a
  category, and for investments: each broker's snapshot and where its return starts, undated ledger rows, deposits
  without an amount, fund and pension balances older than 45 days. Every step is checked against the sheet, never ticked
  by hand, so the list is also the health check. Each step has a button for what to do; buttons can only call a fixed
  list of functions and open a fixed list of tabs. The Setup Wizard ends pointing to it.
- New source file `08_startHere.gs`; the list of source files lives in one place (`tools/bundle.js`), which the test
  harness reads too. Tests: 155, on `src/` and on the single file (mutation-checked: an unscheduled update, ignored
  emails counted).

### v1.1.46 (Sept 28, 2026) — one file to install and update
- The tracker is published as **one file, `dist/FinancialTracker.gs`**: the seven files of `src/` in load order, built
  by `npm run bundle`. Installing and updating is pasting one file — nothing can be half-updated, mixed with another
  version or pasted into the wrong file (what removed the menu in v1.1.33). The whole test suite runs against it too
  (`npm run test:bundle`), and CI fails when `dist/` isn't what `src/` builds (`npm run bundle:check`).
- Considered and not taken: an Apps Script **library**. Per Google's documentation, a library's script properties and
  script lock are one instance shared by every script that includes it — two people's run records and runs would mix
  and block each other — its dialogs can't call library functions directly (each button needs a bridge function in
  every user's project), libraries add latency, and updating still means changing a version by hand.
- The user guide's install steps said six files and missed `07_investments.gs`; they now describe the one file, how to
  update, and how to move from the seven files. `clasp` pushes `dist/`. Tests: 149, on `src/` and on the single file.

### v1.1.45 (Sept 28, 2026) — a deposit notice edited into the deposit clears its flag
- **Reported:** a HAPI deposit notice stayed in Unrecognized after its amount was added — by editing the Notice row itself
  (type changed to Deposit, amount typed). Only a Notice paired with a separate Deposit was cleared, so with no Notice
  left its row stayed forever. A Deposit row carrying the notice's Gmail id now clears it too: both ways of recording the
  amount work. Tests: 147.

### v1.1.44 (Sept 28, 2026) — marking only what's pending; timings in the log; start-day sales are cash
- **Reported:** a year-long range was still killed — reading ended after about a minute, and nothing was logged until
  the 6-minute limit. The time went to marking: every thread in the range was marked read and labelled again, even the
  ones earlier runs had marked — minutes of Gmail calls over hundreds of threads. Gmail now lists the threads still
  unread or without the label (`{is:unread -label:Procesado}`), only those are marked, and marking stops at the time
  budget. The search returns its query for this.
- Every phase logs how long it took (⏱ save, mark, read log, each step), so a slow one shows in View › Executions.
- The missing-deposit check counts the proceeds of a sale made on an account's start day as cash at the start (a sale on
  Dec 31 that settles in January funds January's purchases without a deposit). Found reconciling a real broker history,
  whose cash matched the recorded flows to within cents once its deposits were complete.
- Tests: 146 (mutation-checked: re-marking, start-day sale proceeds).

### v1.1.43 (Sept 28, 2026) — long ranges finish; stopped runs stay light; undated ledger rows are flagged
- **Reported:** a year-long date range stopped reading in time ("54 thread(s) left") and then hit the 6-minute limit
  anyway. Two causes:
  - Emails read without saving anything — statements, promotions, notices, declined — were read again on every run (only
    saved transactions were skipped), so the same range kept re-reading them. Their ids now go to a hidden **Read Emails**
    sheet, tagged with the tracker version (after an update they're read once more, in case the new version reads them
    differently). Emails that failed aren't logged: they're retried.
  - A stopped run still ran every heavy step (sorting, recategorizing with the full formatting, investments with
    GOOGLEFINANCE). A stopped run now saves, marks and records what it read, and leaves those to the run that completes the
    range; and no step starts after 5 minutes. The summary lists what was left ("⏭ Left for the next run: …").
  - The formatting of every sheet ran twice per run (again after the investments step); now only the investment tabs are.
- **Reported:** two manual deposits showed no date. A ledger row without a date can't be placed in time and wasn't
  counted — silently. It's now highlighted in red, and Holdings warns how many rows aren't counted.
- Tests: 142 (mutation-checked: the read log, deferred steps, the 5-minute brake, the undated-row warning; and the
  deposit dialog's date, which the v1.1.40 test didn't check).

### v1.1.42 (Sept 28, 2026) — HAPI deposits missing from the ledger are flagged, one by one
- HAPI's "Deposit Completed" email has no amount, so it was skipped — and a deposit that didn't also arrive as a bank
  transfer the tracker reads went unrecorded, counted as gain. The email is now kept as a **Notice** row in the Investment
  Ledger (it moves no money), and every refresh pairs each notice with a recorded deposit of the same account dated from
  7 days before to 2 days after it (each deposit used once). Each notice left without one goes to **Unrecognized** —
  "Deposit without amount", with its date and a Gmail link — and leaves it once the deposit is added (➕ Add Balance or
  Deposit › A deposit). Notices on or before the account's tracking start are inside its start value and aren't checked.
- Tests: 137 (mutation-checked: the date window, one deposit per notice, notices before the start).

### v1.1.41 (Sept 28, 2026) — the balance/deposit dialog works again; every dialog's script is tested
- **Reported:** in ➕ Add Balance or Deposit, "A balance" and "A deposit" still showed the Units fields. v1.1.40 added a help
  text with an apostrophe ("account's") inside a quoted string of the dialog's own script; the template literal consumed
  its escape, the browser got a broken string and the whole script failed to load — so no option switched and **nothing
  could be saved**. The text no longer has it; the dialog declares UTF-8 and says "Deposit date" for deposits.
- Why the tests missed it: they checked the dialog's HTML for text and never ran its script. A new test compiles every
  dialog's page script as the browser gets it (Setup Wizard, date range, balance/deposit) and checks that every button
  and option calls a function that exists — it fails on the v1.1.40 file. The three modes were also exercised in a real
  browser. Tests: 132.

### v1.1.40 (Sept 28, 2026) — missing deposits are flagged; deposits can be added by hand
- **Reported:** HAPI's 2026 gain looked like it included the money put in. The maths was right (value − start − deposits);
  the data wasn't: only the deposits the tracker saw as bank transfers were recorded, and every missing one shows up as
  gain. Performance now checks each account: when the purchases since its start (net of sales and dividends) exceed the
  recorded deposits by more than US$50 and 5%, it says by how much — unless that came from cash already in the account,
  deposits are missing.
- **📊 Tracker › ➕ Add Balance or Deposit** (was "Add Fund / Pension Balance") gains a third option, *A deposit*: money
  put into an account on a date (a broker funded another way, pension contributions). The same deposit twice is refused.
- Tests: 130 (mutation-checked: the funding check, duplicate deposits).

### v1.1.39 (Sept 28, 2026) — categories you set by hand are kept; HAPI's return from an earlier date
- **Reported:** recategorizing — from the menu and in every daily run — wiped categories set by hand: a transfer you
  categorised went back to empty, a purchase you re-categorised went back to the tracker's choice. And Bank Transfers,
  where transfers are reviewed, is rebuilt from Transactions, so a category typed there was lost too.
  - Transactions gains a hidden **Auto Category** column (O): the category the tracker last set. A category different
    from it was set by hand and is kept. Rows from before this version: a transfer with a category no rule gives is taken
    as yours. Clear a cell to get the automatic category back.
  - Bank Transfers gains a hidden **Id**: before recategorizing, a category typed there is copied to Transactions as yours
    (a sheet from before the Id is matched by date, beneficiary and amount).
- **Reported:** HAPI showed "tracked since 2026-09-26" — an account with positions started at its latest snapshot. It now
  starts at the earliest of its first snapshot and its first balance: add HAPI's value on a past date (➕ Add Fund /
  Pension Balance, kind Broker) and its return runs from there, deposits since then counted as deposits. That balance
  only marks the start — it isn't listed under Other accounts.
- Portfolio History: every other day is banded in a visible colour (the row stripe was too faint to tell days apart).
- Tests: 126 (mutation-checked: manual categories kept, Bank Transfers edits carried back, HAPI's earlier start).

### v1.1.38 (Sept 28, 2026) — HAPI limit orders; Recategorize reorders tabs; hidden tabs stay hidden
- **HAPI limit orders** (from a live email in Unrecognized): a limit order's Cost includes the fee — e.g. 2 × 150.25 + 2.99
  = 303.49 — while a market order's doesn't. Both are accepted now; the cost basis is quantity × price either way (what HAPI
  itself shows as cost), the fee goes in its column and the row is noted "Limit order". A cost matching neither is still
  refused. The quantity was there all along; a one-character quantity line had been dropped while inspecting the email.
- **Reported:** tabs weren't reorganized after 🔁 Recategorize — that menu item never reordered them (only runs, the
  wizard, the Dashboard and investments did). It does now.
- Reordering activates each tab, which can show a hidden one: hidden tabs are hidden again afterwards, and nothing is
  moved when the order is already right. The test mock now treats activating a hidden tab as showing it (the worst
  case), and supports confirmation dialogs (`Button`, `ButtonSet`). Tests: 120.

### v1.1.37 (Sept 28, 2026) — investment changes account by account; distinct status colours; tab groups
- **Reported:** the daily email showed "+US$17,104 since Sat 26 Sep" — the day a fund and a pension were added. The change
  compared totals, so accounts added to the tracker counted as gain. Daily and monthly changes are now measured account
  by account: an account added in the period starts from its first recorded value, and the email says so.
- **Reported:** transfers without a category were highlighted in a colour close to Electricity's chip. A test now measures
  the perceptual distance (CIE ΔE) between every status colour and every category chip and requires ≥ 25 — it also
  caught the unreadable-merchant row colour at ΔE 5.8 from Rent's chip. Both changed (#FDBA74, #FCA5A5).
- **Tabs grouped**: spending, then Holdings / Investment Ledger / Portfolio History together, then settings.
  **📊 Tracker › 🙈 Show / Hide Settings Tabs** hides Investment Accounts, Configuration and Categories (Custom Rules stays:
  it's edited often and the emails link to it). Never automatic, and reordering never unhides them.
- Unrecognized: tracking links are removed from "What the email says". Tests: 116.

### v1.1.36 (Sept 27, 2026) — runs stop in time and resume; saved emails aren't read again
- Reported: a long run (a big date range) hit Apps Script's 6-minute limit — killed mid-step, with the "Saving…" toast
  left on screen and the last steps (recategorize, derived sheets, investments, Unrecognized) never run. A hard kill
  runs no `finally`, so the fix is not to reach it: reading emails — the slow part, one `getPlainBody()` each — stops
  3.5 minutes after the action started, which leaves time to save, mark and run every step. Threads not reached stay
  unread and the summary says to run it again; the next run continues.
- **Already-saved emails are skipped before being read** (their Gmail id is in Transactions). Every run used to read
  them again and drop them as duplicates — even the daily run, whose window re-reads the previous month — so a
  stopped range could keep re-reading the same emails and never finish. Same for HAPI emails already in the ledger.
- Tests: 112 (mutation-checked: saved emails not re-read, the time budget, unread threads not marked processed).

### v1.1.35 (Sept 27, 2026) — the Unrecognized sheet; balances update instead of duplicating
- **Unrecognized** lists every bank or broker email the tracker couldn't read — amount not found, a reading error, a
  HAPI email that didn't validate — and emails saved with an unreadable merchant: date, bank, subject, reason, what the
  email says and a link to it in Gmail. One row per email (a later run updates it and keeps your Status: New or Ignore);
  a row disappears once a later version reads that email. The run summary and the daily email's data-health line say
  how many are waiting. Styled like the other tabs (reason and status chips, ignored rows greyed out).
- **Fund / pension balances**: saving the same account with the same statement date again now updates that balance
  instead of adding a second row. An earlier statement date adds the start of that account's history.
- Tests: 108 (mutation-checked: resolved emails drop off, Status kept, balances not duplicated).

### v1.1.34 (Sept 27, 2026) — investment tabs with the Dashboard's look; Código Cash; one date format
- **Holdings redesigned** like the Dashboard: navy title band, six KPI cards (portfolio value, cost basis, unrealized
  P/L, return since start, dividends, deposited), underlined sections, striped tables with total rows, price-source
  chips, a weight colour scale, and two charts — value over time and allocation by account. The Investment Ledger
  (type and source chips), Portfolio History (day bands, TOTAL rows) and Investment Accounts (kind chips) are styled too.
- **Reported:** balances saved as units × price and as an amount showed their date differently ("9/27/2026" vs
  "2026-09-08") — dates were written as text, which Sheets converts or not depending on the cell. Dates are now real
  dates with one format, in Other accounts and Performance.
- **All accounts** in Performance showed later-starting accounts' start values as "Net deposits"; the return was right,
  the row wasn't. It now shows every account's start value and real deposits only.
- **Código Cash** (POPULAR): a withdrawal has no merchant column, so the status word "Aprobada" was saved as the merchant.
  A status word is never a merchant: these rows read "Código Cash (cash withdrawal)" (category Dining/Other) — new ones at
  parse time, saved ones repaired by recategorize, in a narrow step kept outside `computeRecategorization`, which still
  can never rewrite merchants (v1.1.19's C5 test caught the first attempt).
- Test mock: `hideColumns(column, count)` now hides every column, as Sheets does; style calls keep all their arguments.
  Tests: 103 (mutation-checked: text dates, "Aprobada" as a merchant).

### v1.1.33 (Sept 26, 2026) — progress toasts close; a dialog for fund and pension balances
- Reported from the live sheet: "Refreshing investments..." stayed on screen after the summary. A progress toast with
  no timeout stays until another toast replaces it, and nothing replaced it — nor any progress toast on an error path
  (daily run, date range, Setup Wizard). `withRunLock()` now closes whichever is still open when its action ends,
  successful or not; a test runs the actions and fails if a progress toast is left (mutation-checked).
- **📊 Tracker › ➕ Add Fund / Pension Balance**: a dialog that adds the Valuation row (units × unit price, or a
  balance), lists the account in Investment Accounts and rebuilds Holdings — instead of typing ledger rows by hand.
  A new row per statement; Holdings shows the latest of each account. Tests: 97.

### v1.1.32 (Sept 26, 2026) — portfolio history, real returns, investments in the emails
- **Portfolio History** sheet: every run records each position (quantity, price, value), cash and balance-tracked
  account in US$, plus a TOTAL row — one set per day (a second run the same day replaces it). Values are what
  `GOOGLEFINANCE` computed in Holdings, or quantity × price when a cell hasn't loaded.
- **Returns** per account and for all accounts, since each started being tracked (latest Snapshot, first Valuation, or
  first movement): **Modified Dietz** for the period (deposits weighted by the days they were invested) and
  **XIRR** annualized — only after 180 days, since annualizing a few weeks gives absurd numbers. Holdings shows them in
  a Performance block with a line chart of the total. (Sheets keeps embedded charts on `clear()`, so the rebuild removes
  the old chart instead of stacking a new one every day.)
- **Emails**: the daily summary gains an Investments section (value, gain since the previous day with deposits left
  out, biggest movers, return since tracking began); the monthly one, the month's gain vs deposits, dividends, fees and
  allocation. On/off in the Setup Wizard. Investment percentages use two decimals — daily moves are mostly under 1%.
- Tests: 93 (mutation-checked: chart replacement, time-weighted deposits, deposits not counted as gain, one set of
  history rows per day).

### v1.1.31 (Sept 26, 2026) — crypto prices from Coinbase; sanity check only against recent prices
- Reported from a live sheet: ETHUSD and SHIBUSD showed "last known" — `GOOGLEFINANCE` doesn't price crypto pairs
  (a known limitation). Crypto now comes from Coinbase's public spot price (`/v2/prices/<PAIR>/spot`, no key), one
  batched fetch per refresh; stocks and ETFs stay on `GOOGLEFINANCE`. If Coinbase doesn't answer, the last known
  price is used and labelled.
- Design fix found while doing it: the price sanity check compared a live price with the last known one however old
  it was — after months without trades, a real 50%+ move (NVDA has done it) would have frozen the position at its
  old price. The check now applies only when the last known price is 90 days old or newer.
- The test mock gains `UrlFetchApp` (responses per URL, failures). Tests: 84 (mutation-checked).

### v1.1.30 (Sept 26, 2026) — crypto pairs and tiny prices
- Every crypto pair (6+ letters ending in USD: ETHUSD, SHIBUSD…) is priced as `CURRENCY:<pair>`; before, only ETH and
  BTC were mapped, so SHIBUSD had no live price.
- The last known price went into the Holdings formula rounded to 6 decimals — a price like 0.0000061234 became 0.000006
  (+4% on the position). Now written with 10 significant digits, never in scientific notation. Found while loading a
  real portfolio; covered by a test that goes through the whole Holdings build (the old test called the price
  function directly and missed it).
- Prices under a cent show 8 decimals instead of US$0.00. Tests: 82.

### v1.1.29 (Sept 26, 2026) — investments: ledger, HAPI emails, holdings
- New file **`07_investments.gs`** and three sheets: **Investment Ledger** (every movement), **Holdings** (positions
  valued with `GOOGLEFINANCE`, other accounts, total in US$ and DOP-equivalent) and **Investment Accounts**.
- **HAPI emails** are read in every run, like the bank alerts: executed orders (ticker, quantity, average price,
  cost and the per-order fee — its label is blank in HAPI's template) and dividends (net amount, payment date). An
  order is only saved when its own numbers agree (quantity × average price = cost). "Deposit Completed" has no
  amount, so **deposits come from the bank side**: transfers whose beneficiary contains the account's keyword
  (HAPI: its DR collection account) become Deposit rows — the transfer keeps its own category, e.g. Exclude.
- **Snapshots** make positions exact without history: positions start from each account's latest Snapshot and add
  only later movements, so re-reading older emails never counts a trade twice. Funds and pensions tracked by
  balance use a **Valuation** row (units × unit price, or the amount).
- Prices: a live price more than 50% away from the last known one is treated as a wrong symbol and the last known
  price is used instead ("check symbol" in Price source) — a ticker Google maps to another security can't
  silently misprice a position.
- Menu **📈 Refresh Investments**; the run summary reports the step. Checked end to end against the broker's own
  portfolio screen: per-position P/L and percentages match. Tests: 81 (mutation-checked: order validation,
  snapshot day, de-duplication). The test mock's Gmail search now honours `from:`, as Gmail does.

### v1.1.28 (Sept 26, 2026) — formatting of the data sheets
- **Category colours** (one source: `color` in `DEFAULT_CATEGORIES`): each category shows as a coloured chip in
  Transactions, Bank Transfers, every Raw_ sheet, Custom Rules and Categories. "Exclude" is muted; your own
  Custom Rules categories take colours from a rotation. Every text/background pair is tested for WCAG AA contrast
  (4.5:1) — the test caught the first "Exclude" grey at 4.39:1.
- **Rows that need you** stand out: transfers still without a category (amber), rows whose merchant couldn't be
  read or reversals without their purchase (red). Negative amounts (reversals, refunds) are green; alternate rows
  are striped. All of it is conditional formatting, so it follows the data when rows are recategorized or added.
- **Bank Transfers shows each transfer's Category** (new column D) — the sheet for reviewing transfers didn't show
  which ones still had none. Existing sheets are migrated on the next run; the filter is recreated, since filter
  criteria are kept by column number and one on "Amount" would have landed on "Category".
- Navy, frozen headers on every data sheet; date and amount formats applied to whole columns.
- **Custom Rules**: a dropdown of every category in the Category column (typing a new name still creates one);
  the example row is muted. **Configuration**: keys highlighted, a note that settings are edited in the Setup Wizard.
- The Dashboard is unchanged. Tests: 74 (the mock now records styles, so a preview of each sheet can be rendered).

### v1.1.27 (Sept 25, 2026) — other income, credit cards in the Setup Wizard, wizard closes while saving
- **Other monthly income** (Setup Wizard): an amount in DOP or USD, added in full to net income with no
  deductions — in the Dashboard's income box (its own row), the KPI card, and both summary emails.
- **Credit-card table showed no data** (reported from the live Dashboard): since v1.1.24 the defaults were
  bank names with empty cells, so unless an older table had values, nothing showed. Cards are now set in the
  Setup Wizard — per bank, the card you hold (LAFISE Clásica Mastercard, BANESCO Super Cashback, BHD Mi País,
  or "Other card" with a name and rate), the statement-close day and the payment-due day. Rates and notes
  come from a catalogue of public product facts (`CARD_PRODUCTS`); your cards and days live in Configuration,
  never in the code. The same cards drive the cashback tips, which now name the card ("your BHD Mi País").
  An older Dashboard's typed rows are kept until cards are set in the wizard.
- **Setup Wizard no longer sits frozen while saving:** the window closes after 1.5 s (like Monitor by Date
  Range); the sheet shows progress toasts, rebuilds the Dashboard, and ends with a summary (net income, emails,
  cards) — or the error, which is now always shown even with the window closed. Saving runs under the run lock.
- Found by the new tests: the server rejected a bank marked "No credit card" as an unknown card (the form
  filtered those rows, so it hadn't surfaced yet). Tests: 69.

### v1.1.26 (Sept 25, 2026) — redesigned daily email, new monthly summary
- **Daily email redesign** (shared email kit, table layout + inline styles so Gmail renders it): hero number
  with a status chip, KPI tiles (month so far, month-end pace, daily average), "net income used" vs.
  "month elapsed" bars, last-7-days mini chart, yesterday's purchases, variable categories vs. their usual
  month, transfers to review, recommendation cards by severity, card tips, a data-health line (when the
  morning update ran, how many new transactions, emails it couldn't read — recorded by each monitor run),
  buttons to the Dashboard / Transactions / Custom Rules, and inbox preview text.
- **Day-to-day comparison:** bills and fixed costs (rent, gym, phone, subscriptions, electricity) are left out
  of "vs. your daily average" — a rent or power bill no longer makes a day look like a spike. The total still
  includes them. Categories are flagged "above usual" only from 5% over.
- **Monthly summary** (new file `06_monthlySummary.gs`), sent on the 1st about the previous calendar month:
  spent, left over and savings rate; vs. previous month and usual month; fixed vs. variable; a day-by-day
  heat calendar (day-to-day spending); categories with change vs. previous month; top merchants;
  subscriptions detected (with yearly cost); spend by bank and **cashback left on the table** (purchases paid
  with a card that doesn't reward them, when you have one that does); transfers; year to date;
  recommendations; data health. Enabled in the Setup Wizard; menu **🗓️ Send Monthly Summary Now**.
- Configuration gains `notifyMonthly`; a monthly trigger (`onMonthDay(1)`, same hour as the daily email).
- Both emails were checked rendered in Chromium at desktop and mobile widths. Tests: 62.

### v1.1.25 (Sept 25, 2026) — emoji in the daily summary email
- Every emoji in the summary email arrived as six "�" (subject, recommendation icons, footer), while
  characters like — → − were fine: `GmailApp.sendEmail()` breaks characters outside Unicode's Basic
  Multilingual Plane. The HTML body is now pure ASCII (non-ASCII characters as numeric entities, which
  every client decodes), and the subject and plain-text body carry no emoji. Tests: 55.

### v1.1.24 (Sept 25, 2026) — DR payroll deductions, daily summary email, auto-fit columns, no personal data in code
- **No personal values in the code.** The Setup Wizard no longer ships anyone's email or amounts: a first run
  suggests the email of the Google account opening it, and placeholders are neutral (`you@example.com`,
  `0.00`). Default Dashboard card rows come from the banks you select, with empty rates and dates (what you
  type is kept). Personal names were removed from code comments, the README and tests; tests now check this
  generically (no hard-coded emails besides bank senders, no numeric placeholders, no built-in dates).
- **Dominican payroll deductions (DOP salaries).** The wizard offers *automatic* (ARS/SFS 3.04% and AFP
  2.87% with their 2026 caps of RD$232,230 and RD$464,460; ISR from the DGII scale on salary net of those,
  annualized) with a live preview, or *manual* (ARS, AFP and ISR in RD$). USD keeps a flat tax rate.
  Parameters are per year (`DR_PAYROLL_PARAMS`, 2026 from TSS Res. 01-2025 and DGII DDG-AR1-2026-00001);
  Ley 30-26 changes the ISR scale from 2027 — until that table is added, 2027 uses 2026's and is flagged.
  Automatic values are recomputed on each Dashboard build and daily summary. The Dashboard shows
  "ISR — Impuesto Sobre la Renta (auto)" for DOP and "Income tax (x%)" otherwise, as an amount.
- **Daily summary email** (new file `05_dailySummary.gs`): previous day's spending and top purchases, vs. your
  daily average and month pace, transfers to review, rule-based recommendations, card/cashback tips (only for
  cards in your Dashboard card table), and a link to the spreadsheet. Configured in the wizard (on/off,
  recipient, 7–10 AM, sections); sent from your own Gmail by its own trigger. Menu: **📬 Send Daily Summary Now**.
- **Columns fit their data** after every recategorize (so every monitor run) and every Dashboard build: they
  widen up to 420 px, never shrink below the width you set, and hidden columns are left alone. Dashboard
  column B is sized to its longest category label.
- **Education removed** from the default categories (a Custom Rule with Category "Education" brings it back
  as its own Dashboard row).
- Configuration gains `deductionMode`, `ISR`, `notifyEnabled`, `notifyEmail`, `notifyHour`, `notifySections`;
  older Configurations keep working with defaults. Tests: 54.

### v1.1.23 (Sept 25, 2026) — reversals, bank-id dedup, category fixes, Card Payment → Exclude
- **BHD reversals (real .eml):** a purchase followed seconds later by the same row with Estado
  "Reversada" and no merchant was saved as a second RD$488 "Unknown Merchant" expense. Reversed rows are
  now saved as a **negative** amount that takes the original purchase's merchant and category (paired by
  bank + amount + the table's date/time, or by day against rows already saved), so the two net to zero.
  An unmatched reversal stays uncategorized and is reported in the run summary.
- **Duplicate notices (real case, BANESCO transfer sent twice with identical content):** each bank's own
  id — BANESCO "No. Referencia", BHD "Número de confirmación", LAFISE reference — is stored in a new hidden
  **TxRef** column (N); a second email with the same id is a duplicate, in the same run or later ones.
- **Categories:** spaces around `*` are ignored ("UBER * EATS" → Dining); ALISS / Metro Plaza → Dining
  ("METRO" alone is still Transportation); POLA → Groceries; FARMA, FARM (whole word), ANALISA,
  LABORATORIO, ESPEC MED → Health; UBERBV / UBER BV → Transportation; GOOGLE → Streaming & Subscriptions.
- **Card Payment rows get Category "Exclude"** (parse and recategorize); Categories always lists Exclude.
- **Statement close:** BANESCO and BHD on day 3 (filled in only where it was still "—").
- POPULAR payroll deposit notices are skipped as non-transactional (they were "Could not parse").
- Tests: 43 (reversal thread, duplicate notice, payroll notice, new category cases).

### v1.1.22 (Sept 25, 2026) — "Which card for what": the three cashback programs side by side
- The LAFISE-only reference table is replaced by a matrix of 17 kinds of spending × the three programs
  (LAFISE 10/10, BANESCO SuperCashBack, BHD Visa Mi País) with MCC, each program's rate, minimum and cap,
  and the best card per the terms — including the conditional cases (restaurants, fast food, barbershops:
  LAFISE only once its per-cycle minimum is reached). The recommended card's cell is tinted.
- Card notes rewritten from each program's terms. A note is replaced only while it's still one of the
  old built-in defaults (now also on the named-range path); anything you typed is kept.
- Program rules summarized under the matrix (LAFISE all-or-nothing per category and RD$7,000 total cap;
  BANESCO no minimum, monthly caps, credited at the cut; BHD figures come from the product page only;
  BHD debit purchases earn nothing; all programs pay by MCC, which emails don't include).
- Sources: LAFISE 10/10 terms, BANESCO SuperCashBack reglamento (table 1), BHD Visa Mi País product page.

### v1.1.21 (Sept 25, 2026) — Dashboard redesign + LAFISE 10/10 reference
- **New layout:** title band, month picker by name with a **Current month** option (the Dashboard now
  follows today by itself), KPI cards (net income · spent this month with change vs. previous month ·
  remaining · % of income with a progress bar · transfers to review), "Where the money went" with
  in-cell bars, share and monthly average, income & deductions, fixed vs. variable, spend by bank,
  a month-by-month table with a column chart, and the year as a heat map with the selected month
  highlighted. Title and controls stay frozen on screen.
- **Actual spending, not a projection:** the old "Total Expenses Approx." was fixed-this-month +
  AVERAGE variable. Every month figure is now what was really spent in that month.
- **Inputs in named ranges** (`DASH_MONTH`, `DASH_YEAR`, `RATE_USD`, `RATE_EUR`, `RATE_COP`,
  `RATE_UPDATED`, `DASH_CARDS`), so future layout changes can move cells without losing what you typed.
  The first build carries over the v1.1.19/20 year, rates, rate date and cards; the month starts on
  "Current month" (the old D4 had been a hard-coded 9). `onEdit` finds the rate cells by name.
- **Credit cards:** statement close and payment due columns; the LAFISE note is corrected per the
  10/10 terms — only if it was still the old "Restaurants, Gas, Groceries" default.
- **LAFISE "10/10" reference table:** the 10 categories with cashback, minimum per cycle, maximum back,
  MCC and the closest tracker category, the RD$7,000 monthly cap, and the key rules (all-or-nothing per
  category, credited on the next statement, exclusions). Reference only — nothing is computed from it yet.
- Rebuilds remove old charts, conditional-format rules, validations and merges first, so nothing stacks.
- Tests (35): the strict mock now rejects overlapping merges and values written under a merge, and
  every Dashboard formula is linted (balanced syntax, real Sheets functions only, existing named ranges).

### v1.1.20 (Sept 25, 2026) — DiDi rides → Transportation
- Real POPULAR rows in COP ("DiDi CO Ride", "DL*DIDI RIDES", "DLO*Didi") were landing in the Dining
  fallback — no keyword matched DiDi. Added `DIDI` to Transportation (whole-word match, so the `*`
  variants match too). "DIDI FOOD" still goes to Dining via `FOOD`. Existing rows are fixed by
  "🔁 Recategorize" (also runs automatically at the end of every monitor run).

### v1.1.19 (Sept 25, 2026) — full review: 7 critical fixes, 13 minor fixes, 7 improvements, test suite
**Critical**
- **C1 — last day of every month was never imported.** Gmail's `before:` excludes its date, so the
  daily run skipped today, the run on the 1st searched an empty window, and date-range runs dropped
  their own end date. Now exact epoch bounds (`buildRangeBounds()`), end date included, and the daily
  window starts on the 1st of *yesterday's* month (`computeDailyWindow()`).
- **C2 — concurrent runs duplicated rows and corrupted derived sheets.** New `withRunLock()`
  (LockService) around every monitor/recategorize entry point.
- **C3 — 'ACH' inside "CACHAREPA" typed a real purchase as Transfer**, at parse time and again on
  every recategorize (which passed the merchant as if it were the email body). Whole-word matching for
  short keywords; confirmed consumo subjects pin Type = Transaction; recategorize re-derives Type from
  the subject only and otherwise keeps the saved Type.
- **C4 — a blank Custom Rules keyword matched everything; a numeric one (an account's last 4 digits) crashed every parse.**
  Rules are sanitized; incomplete rows are skipped.
- **C5 — recategorize permanently overwrote merchants that "looked garbled".** Removed; garbled-text
  detection only happens at extraction time now.
- **C6 — `SCRIPT_VERSION` was stuck at 1.1.12.** Now 1.1.19.
- **C7 — rebuilding the Dashboard reset the exchange rates and forced month 9.** D4, G4, G5/I5/K5,
  M5 and the Credit Cards box are preserved; month defaults to the current one.

**Found while implementing (verified against v1.1.18):** a *declined* BANESCO charge was saved as
spending anyway, plus a second fake row with the "Dispones de RD$ …" balance — the extractor dropped
the declined row, so the generic fallback ran. Extractors now flag declined rows, a fully declined
email is skipped without fallback, and balance/limit amounts are never treated as charges.

**Minor**
- M1 whole-email "declined" filter only fires when no row was approved (multi-row BHD tables keep their
  approved rows) · M2 removed generic `RENT`/`RENTA`/`LUZ`, whole-word short keywords · M3 IsCredit
  means money in, scoped to the item · M4 unparseable threads stay unread, "Procesado" label created
  once and applied in batches · M5 paginated search (2,000-thread cap, reported) · M6 messages outside
  the range are skipped · M7 Raw_<BANK> Notes survive rebuilds · M8 Year selector + custom-only
  categories in the grid · M9 month headers at midnight (DATE formulas) · M10 date-range dialog
  validates start ≤ end, closes after 1.5 s, errors are always shown · M11 Setup Wizard pre-fills and
  validates; clear error for a broken banksToTrack; View Config shows currency and version · M12 only
  this system's trigger is deleted/recreated · M13 dead code removed (`getTransactions`, summaries,
  pivot, CSV export, stats, `cleanOldTransactions`, `getUnreadTransactionCount`, rolling-window
  search, `amountPattern`, `getMergedCategoryRules`).

**Improvements**
- E1 batch save: one read + one `setValues()` per run instead of a full-sheet read per transaction ·
  E2 batch recategorize; parse and recategorize categorize on the same text (merchant) · E3 named
  ranges for Configuration, block `setFormulas()`, Google Finance reference row, auto-stamped rate date ·
  E4 extractors declared per bank in `BANK_PATTERNS` · E5 one shared keyword engine · E6 test suite
  (34 tests, real anonymized fixtures, strict mock) · E7 run summary with every counter.

**After deploying:** delete the stale rows (HOLA PLAZA LAS AMERICAS 2026-02-16, the two BHD
"Unknown Merchant" rows, the BHD transfer 2026-09-03) and re-run their dates; then run a date range
over 2026 to pick up the missing month-end days (dedup makes this safe).

### v1.1.18 (Sept 25, 2026) — real BANESCO merchant was getting nuked by its own name; Date Range dialog now closes on Run; vet clinic across all 3 banks
- **"LA CASITA DE" (the vet clinic) wasn't categorizing for POPULAR.** The existing keyword was `'LA CASITA
  DE BOB'`, which matched LAFISE/BHD's full merchant text ("LA CASITA DE BOB-EP SANTO DOMINGO DOM") but not
  POPULAR's, which truncates the merchant name to just "LA CASITA DE" (no "BOB-EP" suffix) — so POPULAR rows
  fell through to the Dining fallback instead. Narrowed the keyword to `'LA CASITA DE'`, a substring of the
  old one, so it now matches all three banks' variants with a single keyword.
- **Real bug found, from a real .eml**: `looksGarbled()`'s `'HOLA '` substring check — added to catch a broken
  fragment ("Hola <name>, Acabas de realiza...") from several versions back — was a false-positive magnet. A
  real BANESCO merchant, **"HOLA PLAZA LAS AMERICAS"** (a real business inside Plaza Las Américas mall),
  extracted correctly by `extractBANESCOConsumoTransactions()`, then got nuked into
  `"(unparsed — see Email Subject)"` by this same check during the very next auto-recategorize pass — which
  runs at the **end of every monitor run**, so a fresh transaction there gets corrupted immediately after
  being saved correctly, not just once. Confirmed the original fragment this rule targeted is still caught
  without it (starts lowercase AND has a bracket — either alone is sufficient). Removed the `'HOLA '` check.
- **Two real "Unknown Merchant" BHD rows, confirmed from real .eml — already fixed by prior versions, not a
  new bug.** Both (`CACHAREPA CHURCHILL` $642, `PedidosYa*Som Cafe` $488) extract correctly against the
  *current* code — the rows are stale, saved by a script version predating the v1.1.11 fix, and never
  reprocessed since. **Fixing a stale row**: delete it from the sheet first, then re-run "📅 Monitor by Date
  Range" for that specific date — re-running without deleting won't help, since `isDuplicate()` keys off the
  email's `messageId` and will just skip it as an existing duplicate rather than re-parsing it. The BHD
  Transfer row confirmed working in v1.1.17 (RD$91.37, "<ACCOUNT HOLDER NAME>") is the same situation — same fix.
- **"📅 Monitor by Date Range" dialog now closes immediately on Run**, per request, instead of sitting open
  for the whole batch. `runGmailMonitorForDateRange()` already reports progress independently via toasts
  ("Searching Gmail...", "Parsing N email(s)...", "Recategorizing...", "Done.") and a final summary alert —
  both are native Sheets UI, not tied to this dialog — so there was nothing the open dialog was adding.

### v1.1.17 (Sept 24, 2026) — LAFISE card payments were double-counting as expenses
- **Real bug found**: the confirmed real LAFISE subject "¡Realizaste un pago a tu tarjeta LAFISE!"
  (active/informal — "you made a payment") never matched the old guessed body-keyword 'PAGO
  REALIZADO A SU TARJETA' (passive/formal — "was made to your card"). These rows got
  Type=Transaction with Category defaulting to Dining (the fallback for an unmatched merchant
  "LAFISE"), meaning paying off the card bill was counted as a brand-new expense on top of the
  original purchases that made up that balance — genuine double-counting. Added the confirmed real
  subject to `TYPE_SUBJECT_KEYWORDS` under Card Payment; these rows now correctly get
  Type=Card Payment (already excluded from the Dashboard, and from Raw_<BANK>/Bank Transfers, by
  existing logic — the only gap was detection).
- **Real bug found**: a live "UBER EATS-W*UBER EATS- SANTO DOMINGO DOM" transaction categorized as
  Transportation instead of Dining. Neither `UBER*EATS` nor `UBEREATS` matched that exact spacing
  (space, no asterisk, between UBER and EATS), so it fell through Dining entirely and landed on
  Transportation's bare `UBER` keyword instead. Added bare `UBER EATS` (space-separated) to Dining.

### v1.1.16 (Sept 24, 2026) — BHD transfers: asterisk-tolerant amount, own-account fallback identifier
- Real detail confirmed from a log: `"Monto: *RD$ 2,950.00"` — an asterisk directly before the
  currency (Markdown-style bold from an HTML→text conversion) that the amount regex didn't account
  for at all. Fixed.
- Rewrote the extractor's structure: finds every `Monto:` occurrence independently, then looks
  **forward** (bounded to before the next `Monto:`) for that transaction's own `Beneficiario:`. If
  none is found there — the working theory, inferred from garbled fragments seen in logs, is that
  BHD's template has no `Beneficiario:` field at all for an **internal transfer between the user's
  own accounts** (no external beneficiary to name) — it looks **backward** instead (bounded to
  after the previous `Monto:`) for that same transaction's own `Producto destino:` account, which
  structurally comes *before* `Monto:`, not after, and uses its last 4 digits as a traceable
  identifier: `"BHD Transfer (to account ...7777)"`. Only falls back to the plain placeholder if
  neither is found.
- Verified against the one confirmed real single-transaction sample (still extracts correctly) and
  a constructed two-transaction bundle (one internal, one to a named beneficiary) — each
  transaction now correctly picks up *its own* destination account, not a neighboring
  transaction's (an earlier version of this idea got this backward and would have grabbed the
  wrong account). **Still not verified against a real bundled multi-transaction email** — the
  field order and the internal-transfer theory are both inferred from log fragments, not confirmed
  from a raw sample. This is a real, structurally-reasoned improvement over the plain placeholder,
  but a genuine `.eml` of a bundled BHD transfer would let it be verified properly instead of
  inferred.

### v1.1.15 (Sept 24, 2026) — the real BHD garble root cause, plus two keyword fixes
- **Real bug found — the garble guard's own rejection was causing the garbled text, not
  preventing it.** Traced from a live log where the guard's target text (`": *DO60BCBH..."`)
  showed up anyway despite the guard demonstrably matching it. Root cause: `extractBHDTransferTransactions()`
  used `continue` to reject a garbled match, which (if it was the only match in the email) left
  `results` EMPTY — and `parseEmailMessage()` treats an empty result as "this extractor found
  nothing at all" and falls back to the GENERIC `extractAllAmounts()`, which has no garble
  protection of its own and re-extracts equally-broken text from scratch. The guard's rejection was
  triggering a *worse*, unprotected fallback path instead of preventing bad data. Fixed: a garbled
  match now keeps the transaction (the amount/date/currency were never the problem) with the
  placeholder text directly, instead of being discarded — so `results` is never empty just because
  of this, and the fallback never fires for this reason anymore.
- **Vehicle Gas keyword fix**: `TOTAL <branch>` (no "ENERGIES" suffix — "TOTAL MIRAMAR", "TOTAL LA
  VECINA") and `NEXT LINCOLN` (a real gas-station brand, confirmed as "NEXT LINCOLN GASOPOLIS" from
  LAFISE and bare "NEXT LINCOLN" from POPULAR) now categorize as Vehicle Gas. Bare `TOTAL` was
  deliberately left out before because "BONJOUR TOTAL ARENOSO" (→ Dining) also contains "TOTAL" —
  solved by reordering Dining/Delivery + Entertainment + Other to be checked *before* Vehicle Gas
  (previously the reverse), so "BONJOUR" intercepts that specific case first. Re-tested every
  previously-confirmed category case alongside the new ones — no regressions.
- **Amazon keyword fix**: bare `AMAZON` used to sit under Streaming & Subscriptions, catching every
  Amazon purchase (retail orders, Marketplace) as if it were the Prime subscription. Narrowed
  Streaming's keyword to `AMAZON PRIME` only; general Amazon purchases now fall under Dining/
  Delivery + Entertainment + Other's bare `AMAZON` instead. Streaming is still checked before
  Dining, so `AMAZON PRIME` specifically still gets intercepted correctly before Dining's broader
  keyword would otherwise also match it.

### v1.1.14 (Sept 23, 2026) — Bank Transfers/Raw_<BANK> could drift out of sync with Transactions
- **Real architectural bug, confirmed from a screenshot**: a row with Email Subject "BHD
  Notificación de Transacciones" (a Consumo email, not a Transfer) was sitting in the Bank
  Transfers sheet with merchant "Unknown Merchant". Cause: recategorize corrected Type/Category/
  Currency by patching Transactions, every `Raw_<BANK>` sheet, and Bank Transfers **independently**
  — so a Type correction made on the Transactions row never relocated the row it left behind in a
  derived sheet (or added it to the one it now belonged in). The three sheets could silently
  diverge from each other.
- **Fix**: Transactions is now cleaned as the single source of truth, and `Raw_<BANK>` + Bank
  Transfers are entirely **rebuilt** from it afterward (`rebuildDerivedSheets()`) — cleared and
  repopulated from Transactions' current Type/Bank/Category/Currency values every time recategorize
  runs (which is now after every monitor run, manual or scheduled). Which sheet a row lives in can
  no longer drift out of sync with its own values, because those derived sheets no longer hold
  independently-patched copies at all — they're always a fresh reflection of Transactions.
  `saveTransaction()` still writes an initial best-guess placement at save time; this rebuild is
  the authoritative reconciliation pass that runs right after.

### v1.1.13 (Sept 23, 2026) — real duplicate-detection bug, DOP income toggle, Transfers now counted everywhere
- **Real data-loss bug, confirmed: two different real transactions could look like duplicates.**
  `isDuplicate()` only compared at DAY granularity — same bank, same day, same exact amount (e.g.
  two separate UBER*EATS orders totaling the same figure) were indistinguishable, and the second
  one got silently discarded. Fixed by keying primarily on the Gmail message's own unique ID
  instead (new hidden `MessageId` column, suffixed per bundled sub-transaction so one email that
  yields several rows — the BHD table case — can't collide with itself either). Falls back to the
  old date+bank+amount heuristic only for legacy rows saved before this column existed. **Legacy
  rows from before this version aren't retroactively fixed** — a same-day-same-amount collision
  that already happened in the past isn't recovered by this change, only prevented going forward.
- **New Setup Wizard currency toggle**: Monthly Income/ARS/AFP can now be entered in DOP instead
  of USD. Stored as `Configuration!B9` (`incomeCurrency`), appended *after* the existing rows
  (not inserted among them) specifically to avoid reintroducing the v1.1.6 wrong-cell-reference
  bug. Dashboard labels are now formulas that read the actual currency instead of a hardcoded
  "(USD)", and "Net Income (DOP equivalent)" only multiplies by the USD rate when the income was
  actually entered in USD.
- **"Are Transfers Not Yet Categorized included in Total Spend / Total Logged Expenses?" — direct
  answer: they weren't, now they are.** It only existed as a Selected-Month-only figure in the
  Income & Taxes box, invisible to the main Category × Month grid and the bottom summary table.
  Added as its own row in the grid (computed for all 12 months, not just the selected one), so it
  now flows into the grid's TOTAL row and from there into "Total Logged Expenses per Month" too.
  Box 2's figure now just reads the selected month from this same grid row instead of duplicating
  the calculation.
- **BHD "Producto destino" still showing in an old log**: re-verified the garble-guard against
  that exact text again — it still correctly flags it as garbled. Given auto-recategorize (v1.1.12)
  now runs after every monitor run, this specific failure mode should self-correct within one
  cycle even in the case where extraction-time somehow still let something through — worth
  checking the `SCRIPT_VERSION` line at the top of a *fresh* run's log (also new in v1.1.12) to
  confirm which code actually produced a given log before treating it as a live discrepancy.
- **Robustness**: `runGmailMonitorCore()`'s parse/save steps are now wrapped in `try/finally`, so
  sort/recategorize/sheet-ordering still run even if something throws a catchable error partway
  through. This does **not** help against a hard execution-time-limit kill — Apps Script
  terminates the whole process in that case, and no code (finally block included) can run after
  that. The only real defense against that remains keeping each run's scope small, which is why
  v1.1.12 narrowed the daily trigger to the current month and "📅 Monitor by Date Range" exists
  for backlogs.

### v1.1.12 (Sept 23, 2026) — the 6:00 AM trigger has been crashing since day one
- **Real bug, confirmed from a real Executions error: the daily trigger never actually worked.**
  `Exception: Cannot call SpreadsheetApp.getUi() from this context.` — a time-triggered execution
  has no UI session at all (no one is looking at the sheet), so `getUi()` throws immediately. Worse,
  the original error handler tried to *report* that failure via another `getUi().alert()` call
  inside the `catch` block, which threw too, uncaught — that second crash is what actually showed
  up in Executions. The real work (search/parse/save) likely completed fine before hitting the
  first alert; only the pointless-with-no-one-watching notification crashed. New `safeAlert()` /
  `safeToast()` wrappers degrade to `Logger.log()` instead of throwing when there's no UI, used
  everywhere the monitor pipeline can be reached from a trigger. Menu-only functions (Setup Wizard,
  Reset System, etc.) still use `getUi()` directly — they're never triggered, always menu-clicked.
- **`runGmailMonitor()` (and the trigger that calls it) now scoped to the current calendar month**,
  not a rolling 60-day window — per request, a function meant to run daily doesn't need 60 days of
  lookback every time, and a narrower window runs faster and further reduces execution-time risk.
  For an actual backlog, "📅 Monitor by Date Range" is still the right tool — that's what it's for.
- **"🔁 Recategorize" now runs automatically at the end of every monitor run** (manual and
  scheduled), per request — Type/Category/Currency corrections apply immediately without a
  separate manual step. The result alert now shows how many cells it touched.
- **`SCRIPT_VERSION` bumped and actually used** — was stuck at `"1.0.0"` since the very first
  version. Now logged at the start of every monitor/recategorize run (View > Logs), so it's
  possible to tell at a glance whether a specific run used the latest deployed code, instead of
  guessing after the fact from symptoms alone (this has come up more than once).
- **LAFISE transfer stopgap improved**: now tries a handful of common Spanish reference-number
  labels (Referencia / No. Referencia / Número de confirmación) and includes whatever it finds —
  "LAFISE Transfer (ref: XXXXX)" instead of the fully generic placeholder — so a row is at least
  traceable back to the real transfer even without a confirmed beneficiary field. Falls back to
  the plain placeholder if none of those labels match.
- **BHD "Producto destino" still showing despite the garble-guard**: re-verified the guard logic
  directly against the exact log text (`": *DO60BCBH0000000000XXXXXXX0011* Producto destino:"`) —
  it matches and should be discarded by the currently-shipped code. The most likely explanation is
  the same one as several times before: the version that produced that specific log predates this
  fix being deployed. Check the new version-logged line at the top of a fresh run's log to confirm,
  then re-run "🔁 Recategorize" to clean up the specific already-saved bad row (it now cleans
  garbled Merchant/Description text too, see v1.1.10).
- **New Dashboard row: "Transfers Not Yet Categorized"** — sum of `Type = Transfer` rows with a
  blank Category (multi-currency, same conversion as the main grid), for the selected month.
  Deliberately excludes any Transfer that already has a real category (e.g. the landlord/Rent case —
  already counted in the main grid's Rent row, so including it here too would double-count) and
  now also feeds into Remaining Balance Approx., which previously ignored uncategorized transfer
  spending entirely.
- **New reserved Custom Rules category: `"Exclude"`** — per request, add a Custom Rules row with
  your own name as the Keyword and `Exclude` as the Category to leave a matching Transfer (a
  self-transfer between your own accounts) out of every Dashboard total, including the new row
  above. No new mechanism needed — it reuses `findCustomRuleOverride()` exactly as-is; `"Exclude"`
  simply isn't one of the 11 real category names, so it never matches a grid row, and it's
  non-blank, so it's automatically excluded from "Transfers Not Yet Categorized" too. A
  documentation row demonstrating this is now added when Custom Rules is first created.
- **Executions stuck showing "Running" forever, after a run keeps going past 6 minutes — this is
  a known Apps Script display quirk, not a bug here.** The platform force-terminates an execution
  that exceeds the time limit, but the Executions transcript UI doesn't always update its status
  cleanly afterward, so it can show "Running" indefinitely even though the run is long dead. This
  is exactly why narrower runs (current-month-only for the daily trigger, "Monitor by Date Range"
  in smaller chunks for backlogs) matter — it's not just about speed, it's avoiding this ambiguous
  failure mode entirely.

### v1.1.11 (Sept 23, 2026) — BHD "Unknown Merchant", finally fixed from a real failing log
- **Root cause, confirmed from a real "Amount not found" log entry**: the BHD consumo regex
  required a literal newline (`\n`) between every field (Fecha/Moneda/Monto/Comercio/Estado/Tipo).
  The failing email — `"08/04/2026 10:20 am US $7.85 LIBERTARIO COFFEE ROASTER Aprobada Compra"` —
  has its fields separated by spaces instead, so the strict-newline pattern never matched at all,
  and the whole transaction silently fell through to the generic fallback ("Unknown Merchant").
  This is also the first confirmed real sample with currency **"US"** rather than "RD" — a foreign
  (USD) purchase on the BHD card.
- **Fix**: every field separator switched from `\n` to `\s+` (matches a newline OR a space), so
  the same regex now handles both layouts. Since `\s+` alone is ambiguous with the spaces *inside*
  a multi-word merchant name (would have stopped the non-greedy merchant capture too early, e.g.
  misreading "LIBERTARIO" as the merchant and "COFFEE"/"ROASTER" as Status/Type), the boundary
  right after the merchant is now anchored to the known status words
  (`Aprobada`/`Rechazada`/`Declinada`) instead of a generic `\w+`. Re-verified against both
  previously-working real samples plus this new one — all three extract correctly with the new
  pattern, no regression.

### v1.1.10 (Sept 23, 2026) — LAFISE transfers never had a real extractor, plus a garbled-data cleanup pass
- **Real bug found: LAFISE transfers were never dispatched to their own extractor at all.**
  `extractTransactionItems()` always called `extractLAFISETransactions()` (the CONSUMO one, which
  requires the "Comercio/Ciudad/País:" label) regardless of Type — a real LAFISE transfer doesn't
  have that label, so it always returned `[]` and silently fell through to the generic
  backward-looking fallback. That's exactly what produced garbled merchants in a live run ("cia a
  banco local] Hola <name>, Acabas de realiza...", "cia cuentas LAFISE] Hola <name>, Acabas de
  realiz...") — mid-sentence fragments, not real data. New `extractLAFISETransferTransactions()`
  is a deliberate **stopgap**: no real LAFISE transfer `.eml` has been seen yet, so it extracts
  only the first amount in the body and labels the merchant generically ("LAFISE Transfer (details
  unconfirmed)") rather than guessing at a beneficiary field whose label isn't confirmed to exist.
  A real sample would let this become a proper extractor like BANESCO/BHD/POPULAR's.
- **New shared `looksGarbled()` detector**, used to both reject a bad capture live and (new) clean
  up already-saved rows. Real bank/beneficiary text is always UPPERCASE in every sample seen so
  far, so a capture starting lowercase is itself treated as a strong garble signal, on top of the
  specific fragment patterns already known (brackets, "Producto", "*DO" + digits, "Hola ",
  "Acabas de"). BHD's transfer extractor now uses this shared check instead of its narrower inline
  one.
- **"🔁 Recategorize Saved Transactions" now also cleans up garbled Merchant/Description text**
  (Transactions, every `Raw_<BANK>` sheet, and Bank Transfers) — replaces an already-saved garbled
  capture with a clear `"(unparsed — see Email Subject)"` placeholder instead of leaving confusing
  fragments in the sheet. Doesn't touch Amount/Date/Bank/Type/Category on that row.
- **Dashboard multi-currency formula reconfirmed correct** — a live formula-bar screenshot showed
  all four `SUMIFS` terms (DOP + USD×rate + EUR×rate + COP×rate) present and correctly referencing
  the exchange-rate cells, matching what shipped in v1.1.8. If a specific EUR/COP transaction still
  looks missing, the likely cause is either it hasn't been picked up by a monitor run yet, or its
  Currency cell holds a stale invalid value from before v1.1.6 — recategorize should now fix the
  latter (see above).
- **How to share a "Amount not found" log for a BHD "Unknown Merchant" row**: open
  [script.google.com](https://script.google.com), open this project, click the clock/history icon
  (**View executions**) in the left sidebar, find the execution around the time that row was
  saved (or re-run "📅 Monitor by Date Range" for that specific date to regenerate it fresh), open
  it, and copy the log lines — specifically the one starting `Amount not found | BHD | ...`, which
  includes a body snippet. Paste that text directly into the chat.

### v1.1.9 (Sept 23, 2026) — POPULAR COP extraction, credit-limit notices, Dashboard now includes categorized Transfers
- **Real bug found: POPULAR's own extractor never accepted COP as a currency prefix.**
  `detectCurrencyFromMatch()` was fixed to *recognize* COP a few versions back, but the actual
  *extraction* regex in `extractPOPULARTransactions()` (and, it turned out, every other bank's
  extractor) still only matched `RD|DOP|USD|US$` — a real COP sample ("COP$33,000.00") never
  matched at all, so the transaction was silently dropped entirely, not just mis-tagged. This is
  what looked like "transactions within a POPULAR thread not being returned": whichever message in
  the thread happened to be COP-denominated failed outright while DOP/USD ones in the same thread
  succeeded. Added EUR/COP to every extractor's currency prefix (LAFISE, BANESCO ×2, BHD ×2,
  POPULAR, the generic fallback), not just POPULAR's.
- **New: credit-limit-increase notifications filtered out.** A real POPULAR "Actualización de
  Límite" email — mentioning the new/old limit as a dollar figure — was getting parsed as if it
  were a purchase (this explains an earlier odd row: 50,000/70,000 DOP under POPULAR with no real
  merchant). Added limit-notice phrases to `NON_TRANSACTIONAL_KEYWORDS`.
- **Dashboard now includes categorized Transfers, not just card Transactions**, per request — a
  Transfer that picked up a real category via a Custom Rule (the landlord/Rent case) should count
  toward that category's total. Changed the grid's Type criterion from an exact match on
  `"Transaction"` to excluding `"Card Payment"` and `"Cashback"` instead. This is safe by
  construction: an uncategorized Transfer (the normal case — internal money movement, Regla 3)
  still has an empty Category and never matches any category row regardless of Type, so this only
  pulls in Transfers that already have a real category assigned.

### v1.1.8 (Sept 23, 2026) — the Dashboard goes multi-currency
- **Same root cause behind two separate complaints: a DOP-only Dashboard next to a USD-only
  income.** Confirmed from a live screenshot — Total Expenses Approx. % showed **2453.6%** and
  Remaining Balance showed **($80,021.44)**, because Monthly Income ($3,400, entered in USD via
  Setup Wizard) was being compared directly against DOP-denominated expenses (~RD$71,050) with no
  conversion. The same DOP-only filter was also why **Streaming & Subscriptions showed "-" every
  month despite real transactions existing** — Spotify, Amazon, and OpenAI are typically billed in
  USD even for a DR cardholder, so a DOP-only `SUMIFS` was correctly, but silently, excluding them.
  Not a date-validation bug — the dates were fine; the currency filter was the actual gap.
- **Fix: the whole Dashboard is now DOP-equivalent, not DOP-only.** New editable exchange-rate
  cells (row 5: USD→DOP, EUR→DOP, COP→DOP — seeded from a live check, ~RD$59.50/USD as of Sept 23,
  2026 via Wise/Xe; EUR/COP are rougher placeholders). Every grid cell now sums all 4 currencies
  (DOP + USD×rate + EUR×rate + COP×rate) instead of filtering to DOP only. Income & Taxes gained a
  new "Net Income (DOP equivalent)" row (USD net income × the USD rate), and that DOP figure — not
  the raw USD one — is what the Remaining Balance / Total Expenses % actually compare against now.
  **These rates aren't fetched automatically and move over time — revisit and update them
  periodically, not just once.**
- Electricity's "wrong amount" turned out to be the same story in reverse: the DOP figure shown
  was probably correct for DOP-only transactions, just missing anything on that category billed in
  another currency, which the multi-currency fix above now includes.

### v1.1.7 (Sept 23, 2026) — the real reason every Dashboard cell read "-"
- **Dashboard grid was comparing icon-prefixed labels against plain category names — never
  matched, for any category, in any month.** Confirmed straight from the formula bar: `$A18` held
  `"💪 Gym + Calisthenics"` (via `getCategoryWithIcon()`), but `Transactions!$F:$F` stores the
  PLAIN name only (`categorizeTransaction()` never adds an icon) — so every `SUMIFS` compared
  `"💪 Gym + Calisthenics"` against `"Gym + Calisthenics"` and got zero, every time. This is what
  the last few "Dashboard not updating" reports were actually hitting, on top of the frozen-column
  crash and the wrong Configuration cell references already fixed. The grid's row labels are now
  the plain category name, matching what's actually stored. (The small Fixed/Variable reference
  table further down the sheet still shows the icon — that one's matched by row *position*, not by
  comparing text, so the icon there was never a functional bug, just cosmetic.)
- **New managed "Categories" sheet** — a reference sheet (Icon / Category / Type / Keywords),
  auto-rebuilt from `DEFAULT_CATEGORIES` every time the Dashboard rebuilds, so it can't go stale
  the way a manually-copied version (e.g. one carried over from the Excel workbook) does when the
  categorizer changes in code. Added at the end of the sheet order, after Configuration.
- Confirmed: BHD's consumption subject is exactly `"BHD Notificación de Transacciones"`, already
  what `BANK_PATTERNS`/comments have on file — no gap there. The still-open BHD extraction issues
  (occasional "Unknown Merchant", transfer-format garbling) are unrelated to the subject line
  itself; still need a real raw sample of the specific failing email to fix precisely.

### v1.1.6 (Sept 23, 2026) — Dashboard crash, wrong cell references, currency over-correction
- **Dashboard crash: "Sorry, you can't freeze columns which contain only part of a merged cell."**
  Real bug in the v1.1.5 Dashboard rebuild — several section headers/labels are merged across
  columns (e.g. "Income & Taxes" spans A6:C6), and `setFrozenColumns(1)` tries to freeze only
  column A, splitting those merges down the middle, which Sheets refuses. Removed the column
  freeze entirely (frozen rows, via `setFrozenRows`, still work and matter more here) rather than
  redesign every merge to avoid column A.
- **Dashboard showing "weird values" — a real off-by-one in the Configuration cell references.**
  The Income & Taxes box's formulas pointed at `'Configuration'!B5:B8`, but Configuration's actual
  layout (row1=header, row2=email, row3=monthlyIncome, row4=ARS, row5=AFP, row6=taxRate) puts
  those values at `B3:B6` — two rows off. Fixed. Because the crash above meant the function never
  even finished a full run before, this bug hadn't been visible yet either.
- **Currency detection over-corrected — real bug from the v1.1.5 broadening.** Widening currency
  detection to "any 2-3 letter code" (to catch a real COP transaction) was too permissive: a live
  run showed it capturing garbage — "FKR", "WEC", "QAH", "JDH", "UCO", "BGJ" — random uppercase
  fragments from reference numbers and authorization codes elsewhere in the email that happened to
  precede a digit. Reverted to an explicit fixed list per request: DOP, USD, EUR, COP only (plus
  RD→DOP, US→USD). `recategorizeAllTransactions()` now also resets any already-saved Currency
  value outside that set back to DOP, to clean up rows saved during the brief window the bad
  version was live.
- **`getUserCustomRules()` email matching now trims whitespace and ignores case** — a mismatched
  email between the Custom Rules sheet and Configuration would otherwise silently return zero
  rules, which is one way the landlord/Rent override could fail even with the right keyword. If the rent transfer
  is still uncategorized after this: confirm the Custom Rules row has `UserEmail` exactly matching
  your Configuration email, `Category` = `Rent`, and `Keyword` = the landlord's first name (not the accented full
  name — see the v1.1.5 note on why), then run "🔁 Recategorize Saved Transactions" again.
- **IsCashback column hidden** too now (same treatment as Timestamp) — "no necesito mostrar si es
  cashback o no." Still written on every save, just not shown by default.
- **Setup Wizard now lands on Dashboard if one already exists** (re-running setup on an existing
  system), otherwise Configuration (first-time setup — Dashboard has nothing to show yet).
- **BHD consumption extractor re-verified against a second real sample** ("SOLO PASTELITOS
  GUILBE") — matched correctly, so the earlier "Unknown Merchant" row was likely a one-off (a
  different email subtype, or a Gmail rendering quirk) rather than a pattern bug — no evidence yet
  to fix further without seeing that specific email's actual text.
- **BHD transfer garbling**: no new sample was provided this round (the attached .eml was the
  consumo type, not transfer) — the v1.1.5 sanity-check guard (discards an obviously-broken
  merchant capture rather than saving it) should already reduce bad rows once this version is
  deployed, but the underlying multi-transaction layout is still unconfirmed. A real BHD transfer
  email (ideally one with 2+ transactions the same day) would let this get fixed properly instead
  of just guarded against.
- **If category names in the Dashboard/sheets still show old ones (Internet, Mobile Data,
  Clothing, Other)**: this isn't a code bug on its own — `getCategories()` already returns the
  current 11-category list once `02_categorizer.gs` is deployed. Make sure **all 4 files** were
  replaced completely in the Apps Script editor (not just the ones that changed most recently),
  save, and reload the Google Sheets tab so `onOpen()` rebuilds the menu with current code — the
  same reload step mentioned for the earlier "function not found" issue applies to stale category
  names too.

### v1.1.5 (Sept 23, 2026) — Dashboard rebuild, timezone bug, custom-rule override, currency, and a dozen smaller fixes
- **Dashboard button "deleting everything" — root cause: there was nothing else to delete.**
  `buildOrRefreshDashboard()` only ever built the Category × Month grid; the Income & Taxes /
  Selected Month Summary / Credit Cards boxes were only ever built in the Excel companion, never
  in the live Apps Script version. Rebuilt to include all four boxes, with Income & Taxes wired
  as **live formulas against the Configuration sheet** (`='Configuration'!B5` etc.) so editing
  Configuration updates the Dashboard automatically.
- **Dashboard read Dec-25→Nov-26 instead of Jan-26→Dec-26 — a real timezone bug.** Month headers
  were built at midnight (`new Date(year,m,1)`); if the Apps Script project's time zone (Project
  Settings) differs from the spreadsheet's own time zone (File > Settings) — a very common
  mismatch — midnight in one can land on the previous calendar day in the other. Same root cause
  as the earlier "21:00:00" bug, just hitting month headers this time. Now builds at **noon**
  instead, which absorbs any reasonable mismatch. `formatDate()` (actual transaction dates) got
  the same fix. The permanent fix is aligning the two time zone settings to match; this is a
  defensive floor under that.
- **Custom Rules can now override a Transfer's "stays uncategorized" default.** New
  `findCustomRuleOverride()` (02_categorizer.gs) is checked FIRST, regardless of Type — this is
  what lets a landlord's name categorize a rent transfer. **Important:** if your beneficiary name
  has an accented character (ñ, é, etc.), check what actually landed in the sheet — email
  encoding can mangle it (a real example: a name with an Ñ was stored with the Ñ
  replaced by a space). Use the shorter, unaccented fragment that's actually stored (e.g. just
  the first name) as the Custom Rules keyword, not the fully-accented name.
- **Currency detection broadened beyond USD/DOP.** A real POPULAR sample showed a transaction in
  Colombian Pesos ("COP$587,720.00") that the old USD-vs-DOP-only check missed entirely,
  defaulting it to DOP and risking it silently inflating DOP totals. `detectCurrencyFromMatch()`
  now captures whichever 2-3 letter code actually precedes the amount (DOP, USD, COP, EUR, ...).
- **LAFISE transfers weren't found — a second real sender confirmed.** LAFISE sends consumption
  alerts from `notificaciones@bancolafise.com` but transfers ("<Name>, ¡Transferencia exitosa!")
  from a *different* address, `digital@notificaciones.lafise.com`. `fromDomain`/`searchQuery`
  broadened to match both.
- **`IsCredit` accent bug fixed** — checked for `'CREDITO'` (no accent), but real bank text says
  "Tarjeta de **Crédito**" (accented); `.toUpperCase()` doesn't strip accents, so this never
  matched and IsCredit was always NO. Checks both forms now. `IsCashback` stays NO for now — no
  real "you received cashback" sample exists yet to build real detection from.
- **Sheet order swapped**: Raw_BDI now before Raw_POPULAR, per request.
- **Tab colors added** — every `Raw_<BANK>` sheet shares one common gray; everything else gets
  its own color.
- **"Monitor by Date Range" landing on Configuration — a real side-effect bug.**
  `ensureSheetOrder()` repositions sheets via `setActiveSheet()`+`moveActiveSheet()`, which leaves
  the *last* sheet processed (Configuration) active as an unintended side effect. Now restores
  whatever was active before reordering, and callers that want to land somewhere specific
  (Dashboard, after a monitor run) activate it themselves afterward.
- **Sheets now sort newest-first** (`sortAllDataSheets()`) — Transactions, Bank Transfers, and
  every `Raw_<BANK>` — at the end of every monitor run.
- **Friendlier processing feedback** — a non-blocking toast ("Searching Gmail...", "Parsing N
  email(s)...") shows immediately instead of the person staring at a blank sheet until the final
  alert.
- **Category list consolidated further**: `Housing` → `Rent`; `Internet` + `Mobile Data` merged
  into `Telecommunications`; `Clothing` and `Other` removed as separate categories — both folded
  into `Dining/Delivery + Entertainment + Other` (that's now also the fallback for anything
  unmatched, replacing the old bare `Other`). Added real merchants that were falling through:
  PedidosYa (a delivery app), Krispy Kreme, IKEA Restaurant.
- **BDI added to the Setup Wizard** — was missing entirely; no way to select it before.
- **Trigger time changed from 8:00 AM to 6:00 AM.** If you set up before this change, your
  trigger still fires at 8:00 AM — re-run Setup Wizard (safe to repeat; it always deletes and
  recreates the trigger) to pick up the new time.
- **Timestamp column hidden** in Transactions (still written, just not shown — unhide via Format
  menu if you want it back).
- **Raw_<BANK> sheets now only receive `Type = Transaction` rows** — Transfers still go to
  Transactions (complete ledger) and Bank Transfers (dedicated view), not into the per-bank raw
  sheets anymore, per request.
- **BHD transfer extractor hardened against garbled multi-transaction captures** — a real
  multi-transaction BHD transfer email (confirmed from a live screenshot) broke the merchant
  capture, producing fragments like `": *DO60BCBH..."` instead of a real name. Added a sanity
  check that discards an obviously-broken capture instead of saving it — loses those rows rather
  than mislabeling them, until a real multi-transaction sample is available to fix the pattern
  properly.
- **Not fixed / needs more information:**
  - A BHD consumption row showing merchant = "Unknown Merchant" means `extractBHDConsumoTransactions()`
    didn't match that specific email's structure and it fell through to the generic fallback.
    Check that row's Email Subject and View > Logs for the "Amount not found" line with a body
    snippet, or share the raw text, to fix precisely rather than guess further.
  - Couldn't conclusively identify a formatting change specifically after row ~220 from the
    screenshots — every row gets the same number-format calls in `saveTransaction()` regardless
    of row number, so nothing in the code should behave differently past a specific row. If it
    persists, a screenshot comparing a "before" and "after" row side by side (same columns
    visible) would help pin it down.
  - The "Invalid: Input must fall within specified range" error isn't caused by this codebase —
    nothing here ever sets data validation on the Category column (only the Dashboard's month
    selector, D4). It's almost certainly a leftover validation list from pasting/importing data
    from the Excel workbook, now stale given the category renames above. Data > Data validation >
    select the column > Remove validation clears it.

### v1.1.4 (Sept 23, 2026) — SUPERCASHBACK bug, full English, sheet ordering, fresh Excel
- **Real bug found and fixed: every BANESCO purchase was mistyped as Cashback.** The card
  product itself is named "VISA CLASICA SUPERCASHBACK" (confirmed from a real email) — the bare
  `'CASHBACK'` keyword in `TYPE_KEYWORDS` matched that product name as a substring, so every
  purchase on that card got Type=Cashback instead of Transaction, and since Category is only
  assigned when Type=Transaction, those rows silently got no category — and stayed that way even
  after recategorizing, because recategorization also respected the (wrong) stored Type. Removed
  the bare keyword; kept only more specific reward-language ones (no real "you got cashback"
  sample exists yet, so this Type is best-effort until one does). `recategorizeAllTransactions()`
  now ALSO re-derives Type (from stored Subject + Merchant) and corrects it, not just Category —
  this is what actually fixes already-saved rows; re-run "🔁 Recategorize Saved Transactions".
- **Whole system translated to English** — was a mix (English menus/headers next to Spanish
  sheet names, category names, and Type values). Renamed: `Transacción→Transaction`,
  `Transferencia→Transfer`, `Pago TC→Card Payment`; all 13 categories (Vivienda→Housing,
  Gasolina Vehículo→Vehicle Gas, etc. — see the categorizer's header comment for the full
  mapping); sheet names `Config→Configuration`, `CustomRules→Custom Rules`,
  `Transferencias→Bank Transfers`; every `Raw_<BANK>` and Bank Transfers header
  (Fecha/Tipo/Comercio/Monto/Moneda→Date/Type/Merchant/Amount/Currency); every menu item, dialog,
  and alert message. Keywords themselves stay in Spanish — they match Spanish bank-email text,
  which didn't change.
- **New `ensureSheetOrder()`** arranges sheets as requested: Dashboard, Transactions, Bank
  Transfers, Raw_LAFISE, Raw_BANESCO, Raw_BHD, Raw_POPULAR, Raw_BDI, Custom Rules, Configuration.
  Called after `saveSetupConfig()`, `buildOrRefreshDashboard()`, and every `runGmailMonitorCore()`
  run, so it stays correct as new sheets get created.
- **"Script function not found" isn't a code bug** — Apps Script menus bind function names at
  the moment `onOpen()` runs; if the code was updated but the Sheets tab wasn't reloaded, the menu
  still points at the old (now-removed) function name. Paste the latest code, then reload the
  Sheets tab (or close/reopen it) so `onOpen()` rebuilds the menu.
- **Fresh `Financial_Tracker_2026_v2.xlsx`** — rebuilt from scratch to mirror the current live
  system exactly: same sheet names/order, same English headers, same 13 categories, same Type
  values, live SUMIFS Dashboard. Useful as a clean starting point or as a structural reference.

### v1.1.3 (Sept 23, 2026) — real dates, retroactive recategorization, reset/trigger bugs, date-range dialog
- **"21:00:00" on every row, fixed.** The Date column now gets an explicit `yyyy-MM-dd` number
  format (Transactions, every `Raw_<BANK>` sheet, Transferencias) — the underlying value was
  already a real `Date` object since v1.1.2, but with no format applied Sheets defaulted to a
  datetime display, and since every row's time-of-day was midnight in the script's timezone, it
  rendered as the same constant (wrong-looking) time everywhere.
- **CustomRules now actually used.** The sheet, `getUserCustomRules()`, and `getMergedCategoryRules()`
  existed since v1.0.3, but the live pipeline never called them — `categorizeTransaction()` was
  always invoked with default rules only. `extractTransactionsFromThreads()` → `parseEmailMessage()`
  now thread an optional `customRules` argument through, and `runGmailMonitorCore()` fetches
  `getMergedCategoryRules(config.email)` before each run. For "<LANDLORD FULL NAME>" (or
  similar personal-context rules, like a landlord's name → Vivienda): add a row to the
  `CustomRules` sheet (`UserEmail | Category | Keyword | Timestamp`) rather than hardcoding a
  private individual's name into the shared categorizer — it now takes effect automatically.
- **New "🔁 Re-categorizar Transacciones Guardadas" menu item** (`recategorizeAllTransactions()`).
  `saveTransaction()` writes a static Category value, not a live formula — fixing a categorizer
  keyword only affects rows saved *after* the fix. This re-applies `categorizeTransaction()`
  (CustomRules included) to every already-saved row in Transactions and every `Raw_<BANK>` sheet
  and updates the Category cell in place, so past fixes (BRAVO, FCIA, EDESUR, etc.) apply
  retroactively without re-running the Gmail search.
- **More real merchants**: `HIPER OLE` (was `HIPERMERCADOS OLE` — same lesson as FCIA/MEDICAR,
  the real card text abbreviates it), `365 EL CACIQUE` (confirmed real chain, Supermercado 365 —
  kept as the full phrase rather than bare `365`, a 3-digit number far too collision-prone),
  `HOLA PLAZA` (confirmed real chain, Hola Market — kept as a 2-word phrase rather than bare
  `HOLA`, which is literally the greeting BANESCO's own transfer emails open with: "Hola ,"; a
  bare match would have wrongly categorized every BANESCO transfer as a supermarket purchase —
  verified this doesn't happen), `LA CASITA DE BOB` → Salud/Vet.
- **Reset System now actually resets.** Used to delete only Config and Transactions — every
  `Raw_<BANK>` sheet, Transferencias, Dashboard, the Pivot sheet, monthly Summary sheets,
  CustomRules, and the daily trigger were all left behind. Now discovers and removes everything
  the system creates, and says so in the confirmation dialog.
- **Setup Wizard's "will start automatically" claim, now true.** `createTrigger()` (an 8:00 AM
  daily trigger) existed but was never called from `saveSetupConfig()` — the dialog promised
  automatic monitoring that never actually got scheduled. Now calls it, and the success message
  says when (8:00 AM daily) instead of a vague "automatically."
- **Date-range dialog replaces the two sequential prompts** — both dates are entered together via
  native `<input type="date">` fields (the browser's own calendar picker), pre-filled with last
  calendar month as a one-click default for the common "catch up" case, fully editable for any
  custom range. Kept **both** "🔄 Monitor Gmail Now" (quick incremental runs — also what the new
  daily trigger calls) and "📅 Monitor por Rango de Fechas" (deliberate backlog processing) rather
  than replacing one with the other: the rolling-window and fixed-range use cases are different
  enough that collapsing them into one "last month" button would serve neither well.
- **Raw_POPULAR sheet moved to the last position** on first creation.
- **Dashboard menu item kept, on purpose.** Since `buildOrRefreshDashboard()` (v1.1.2) writes
  live `SUMIFS` formulas, the sheet updates on its own as new rows arrive — the menu item's job
  isn't "refresh," it's create-if-missing and jump-to. Still useful for first-time setup and for
  rebuilding if the sheet is ever deleted by accident.

### v1.1.2 (Sept 23, 2026) — date-range runs, real Dashboard, two real UI bugs fixed, column reorder
- **500-thread question answered:** yes, with a year's worth of notifications across 4-5 banks,
  you'll want to split the backlog into a few date-range runs rather than one pass over all of
  2026 — not just because of the search cap (which now paginates up to `limit`, default 500), but
  because Apps Script triggers have a hard 6-minute execution ceiling regardless. A few hundred
  emails per run is a safer target than a few thousand in one go.
- **New "📅 Monitor por Rango de Fechas" menu item** — prompts for a start/end date
  (AAAA-MM-DD) and runs the same pipeline scoped to that window via
  `searchTransactionEmailsByDateRange()`. `isDuplicate()` still guards every save, so overlapping
  ranges across multiple runs (e.g. re-running January after already having done Q1) never
  double-count — split the year however's convenient.
- **Currency moved next to Amount** in the Transactions sheet (was appended at the very end,
  column L) — now column E, right after Amount (D). This shifted every column after it; all
  reader functions (`getTransactions`, `getSummaryByCategory`/`byBank`, the pivot, stats) were
  updated to a shared `TX_COL` index map instead of magic numbers, so this kind of drift is easier
  to catch next time. Also fixed the filter only covering 11 of the now-12 columns (Currency was
  outside it), and added `#,##0.00` number formatting to the Amount/Monto columns everywhere
  (Transactions, every `Raw_<BANK>` sheet, Transferencias) — they were plain unformatted numbers
  before.
- **Setup Wizard now actually closes.** Real bug: the dialog had two *separate*
  `google.script.run` chains — one called `saveSetupConfig(config)` with no handlers attached, the
  other registered `withSuccessHandler`/`withFailureHandler` (including the `google.script.host.close()`
  call) but never actually invoked a function on that chain, so those handlers never ran. Merged
  into one correctly-chained call.
- **Dashboard now exists and auto-updates.** `openDashboard()` used to just show a "coming in
  Phase 2" alert — there was no dashboard to update, which is why it "wasn't updating." New
  `buildOrRefreshDashboard()` builds a real "Dashboard" sheet with live `SUMIFS` formulas (a
  Categoría × Mes grid, DOP + Type=Transacción only, mirroring the Excel workbook's design) against
  the Transactions sheet. Because these are real spreadsheet formulas, once built the sheet
  recalculates on its own as new rows come in — no re-run needed.
  - This depends on the Date column holding a real date value, which surfaced a second real bug:
    `formatDate()` (`03_gmailMonitor.gs`) returned a locale-formatted **string**
    (`toLocaleDateString('es-DO')`) instead of a Date object — Sheets doesn't reliably
    auto-convert that to a real date, which would have silently broken every date-range formula.
    Now returns an actual `Date` object. Rows saved before this fix may still be text-dated; see
    the code comment for how to spot and fix them if the Dashboard looks like it's missing older
    data.
  - This also meant `isDuplicate()`'s date check (`data[i][0] === transaction.date`, a strict
    reference comparison) would have started silently failing for every row once dates became
    real `Date` objects — two distinct Date objects for the same day are never `===` in
    JavaScript. Rewrote to normalize both sides to a `yyyy-MM-dd` string before comparing.
- **New "Transferencias" sheet** — `Type = "Transferencia"` rows now also get written to a
  dedicated sheet (Fecha/Banco/Beneficiario/Monto/Moneda/Asunto), alongside their existing spot in
  the master Transactions sheet and their `Raw_<BANK>` sheet, for a clean transfers-only view.
- **More real merchants categorized**, several from live transaction data: `AMAZON`, `OPENAI` →
  Streaming & Suscripciones; `SM NACIONAL` → Supermercado (Bravo's two branches already matched
  via the existing bare `BRAVO` keyword); `COMEDOR` → Dining; bare `CLARO` → Datos Móviles (so
  "CLARO P REC" matches — Internet's more specific `CLARO INTERNET` is checked first and still
  wins for actual internet bills); `TOTALENERGIES` (one word, the fuel-pump descriptor) →
  Gasolina — deliberately *not* bare `TOTAL`, since "BONJOUR TOTAL ARENOSO" (the convenience-store
  purchase) contains "TOTAL" as a separate word and Gasolina is checked before Dining; `WENDYS`,
  `JADE TERIYAKI`, `SWEETFROG`/`SWEET FROG` → Dining.

### v1.1.1 (Sept 23, 2026) — search cap, per-bank sheets, filters, categorization fixes
- **Fixed the 50-thread cap.** `searchTransactionEmails()`'s call site in `runGmailMonitor()`
  (`01_main.gs`) was hardcoded to `50` — every run silently dropped anything beyond the 50 most
  recent matching threads, even with a 60-day search window. The function now paginates in
  batches of 500 (Gmail's practical per-call cap) up to a configurable `limit` (default 500), and
  logs when it hits that cap so you know to narrow `daysBack` or raise it.
- **Per-bank Google Sheets now actually populate.** `saveTransaction()` previously only wrote to
  the single master `Transactions` sheet — the `Raw_<BANK>` per-bank tabs (Fecha/Tipo/Comercio-
  Descripción/Categoría/Monto/Moneda/Notas/Asunto del Correo, mirroring the Excel workbook) were
  never created or written to by the live Apps Script system at all. `saveTransaction()` now
  writes to both.
- **Filters added** to both the master `Transactions` sheet and every `Raw_<BANK>` sheet via a
  new `ensureAutoFilter()` helper (range set generously to 2000 rows so appended rows stay under
  it without resetting).
- **New `Currency` column** (master sheet column L, `Raw_<BANK>` column F) — a real POPULAR
  sample showed a transaction in USD ("US$14.27, Dólar estadounidense"), so DOP could no longer
  be assumed for every row. `detectCurrencyFromMatch()` checks each extractor's matched text for
  `US$`/`USD`, defaulting to DOP.
- **Fixed the description field leaking the raw field label.** `description` was built from
  `item.context.substring(0,150)` — the full regex match text, which for LAFISE starts literally
  with `"Comercio/Ciudad/País:"` (the label itself). Now uses the already-clean `item.merchant`.
- **Categorization fixes**, several driven by real transaction data:
  - `COLMADO` removed from Supermercado — a colmado (corner store) isn't a supermarket; falls
    through to Otros now.
  - `'UBER EATS'` (space) → `'UBER*EATS'` (asterisk) — that's the real merchant descriptor format
    seen in a live LAFISE email, so the space version never matched anything. Added
    `UBER*RIDES`/`UBER*TRIP` to Transporte alongside the existing generic `UBER` catch-all;
    Dining is checked before Transporte so the more specific UBER*EATS match wins correctly.
  - Added `HELADOS BON` and `BONJOUR` (the convenience-store brand inside Total gas
    stations — sells sandwiches/coffee, confirmed via search, not fuel) to Dining, per direct
    correction.
  - Added `FCIA` — the actual abbreviation Dominican banks use for "Farmacia" in transaction text
    (confirmed from a real BHD email: "FCIA MEDICAR GBC 30 DE MA" — the old rules only had the
    unabbreviated `FARMACIA`, so this real pharmacy transaction would have fallen through to
    Otros). Added major pharmacy chains researched for this update: Medicar GBC, Los Hidalgos,
    Farmax/FarmaXtra, Cruz Verde, Farmatodo, FarmaValue.
  - Added major supermarket chains beyond Bravo/Carrefour/Jumbo: Supermercados Nacional, La
    Sirena, Plaza Lama, Iberia, Aprezio, La Fuente, PriceSmart, Hipermercados Olé.
  - Added a few major fast-food chains to Dining: McDonald's, Burger King, KFC, Pizza Hut,
    Subway, Pollos Victorina.
  - `BRAVO` (Supermercados Bravo) was already correctly mapped — verified, no change needed.

### v1.1.0 (Sept 23, 2026) — extraction rewritten against 6 real .eml samples
- The user shared actual `.eml` files for all 6 confirmed email types (LAFISE/BANESCO/BHD
  consumo + transferencia, POPULAR consumo). Extraction is no longer proximity-based guessing —
  each bank+type combination now has its own regex, written and unit-tested against the real
  body text before shipping.
- **Key discovery: currency prefix varies by bank/template**, not just the `$` presence fixed in
  v1.0.8. Real formats seen: `RD$ 2,640.00`, `RD 8,315.20` (no $), `DOP 412.86` (LAFISE, BANESCO
  transfers), `DOP545.00` (no space), `RD$ 91.37`, `RD` + `$118.40` split across a table
  cell/newline (BHD), `US$14.27` (POPULAR). `BANK_PATTERNS.amountPattern` (kept for BDI) and every
  new extractor use `\b(?:RD|DOP|USD|US\$)` followed by an *independently* optional `\$?` — that's
  what makes both the attached and cell-split cases match.
- **Key discovery: merchant position relative to the amount is bank-specific.** LAFISE states the
  merchant *before* the amount ("Comercio/Ciudad/País: X ... Monto: Y"); BANESCO, BHD, and POPULAR
  all state it *after* ("consumo de RD$Y, en X y su estado..." / a Monto-then-Comercio table row).
  The old v1.0.9 "context window before the amount" approach was structurally wrong for 3 of the 4
  banks — replaced by six dedicated extractors:
  - `extractLAFISETransactions` — merchant before amount
  - `extractBANESCOConsumoTransactions` — amount before merchant, skips a declined row via its own
    status check (defense-in-depth alongside the whole-email `isDeclinedTransactionEmail`)
  - `extractBANESCOTransferTransactions` — "Nombre del Beneficiario" as merchant
  - `extractBHDConsumoTransactions` — real 6-column table (Fecha/Moneda/Monto/Comercio/Estado/Tipo);
    uses a global regex so it naturally picks up as many transaction rows as exist — this is where
    the same-day-bundling case is most likely to actually surface
  - `extractBHDTransferTransactions` — "Beneficiario" as merchant
  - `extractPOPULARTransactions` — compact one-row table, lower confidence (tab/space layout
    inferred from a single sample)
  - `extractAllAmounts` (renamed scope) is now BDI-only + a safety-net fallback if a specific
    extractor finds nothing
- All 6 real samples were also re-checked against `PROMOTIONAL_KEYWORDS`, `NON_TRANSACTIONAL_KEYWORDS`,
  and `DECLINED_KEYWORDS` — zero false hits, confirming none of the earlier filter tightening
  accidentally excludes a real transaction.

### v1.0.9 (Sept 23, 2026) — confirmed senders/subjects + multi-transaction support
- **All 5 bank senders now confirmed directly by the user**, no more guessing:
  - LAFISE — `notificaciones@bancolafise.com` (unchanged; transfer subject now known: "*, ¡Transferencia exitosa!")
  - BANESCO — `notificaciones@banesco.com.do` (unchanged; transfer subject confirmed: "Notificación de Transferencia Realizada")
  - BHD — **`Alertas@bhd.com.do`** (was a domain-wide provisional rule; now pinned to the exact
    confirmed address, same as the others). Consumption subject: "BHD Notificación de
    Transacciones". Transfer subject: "Transacciones entre mis productos".
  - POPULAR — **`notificaciones@popularenlinea.com`** — re-enabled (was disabled since every
    earlier hit was a false positive). Consumption subject: "Notificación de Consumo".
- **New `TYPE_SUBJECT_KEYWORDS`**, checked before the body-keyword fallback in
  `detectTransactionType()` — subject lines are a far more reliable Type signal than guessing
  body phrasing, and we now have the real ones for LAFISE/BANESCO/BHD transfers.
- **Multi-transaction support**: the user confirmed a bank can bundle several same-day
  transactions into one thread/message. `parseEmailMessage()` now returns an **array** (was a
  single object), and a new `extractAllAmounts()` finds every currency figure in the body — not
  just the first — pairing each with its own nearby context for merchant extraction
  (`extractMerchantFromContext()`, scoped to that context window instead of the whole email).
  `extractTransactionsFromThreads()` flattens the arrays. Best-effort: the exact bundled-email
  layout hasn't been seen in a real sample yet, so this is unconfirmed — every parse still logs
  its context, so the real layout will surface from the next live run.
- Removed the now-dead whole-body `extractAmount()` / `extractMerchant()` functions, replaced by
  the per-item versions above.

### v1.0.8 (Sept 23, 2026) — real amount/merchant extraction fixes from live data
- **Root cause of most "Amount not found" failures:** `amountPattern` used `.*?` between the
  keyword (monto/consumo/pago) and the currency symbol, but JavaScript's `.` never matches a
  newline without the `s` (dotAll) flag. Real bank emails are multi-paragraph, so the keyword and
  the actual `RD$` figure are almost always on different lines — added `s` to every
  `amountPattern` and the generic fallback in `extractAmount()`.
- **"RD 8,315.20" has no `$`** — a real Banesco decline notice confirmed the bank sometimes
  omits the currency symbol entirely. `RD\$` → `RD\$?` (and same for `USD`) everywhere.
- **Added `\b` before RD/USD** — bare "RD" without a word boundary risks matching inside
  ordinary Spanish words that contain that letter pair (tarde, orden, acuerdo, recordar).
- **New `isDeclinedTransactionEmail()` / `DECLINED_KEYWORDS`** — a real Banesco sample was a
  *declined* charge ("ha sido rechazada... TARJETA VENCIDA") with a real amount and merchant in
  the body. No money moved, so it's excluded before ever reaching `extractAmount()`.
- **`isAccountNoticeEmail()` renamed to `isNonTransactionalEmail()`** and `ACCOUNT_NOTICE_KEYWORDS`
  to `NON_TRANSACTIONAL_KEYWORDS`, widened with real categories seen in the same run: account
  statements ("Estado de cuenta..."), fund reports ("Fondo de Inversión"), satisfaction surveys
  ("Queremos conocer tu opinión"), and event invites ("Acompáñanos en...") — none of these are a
  single transaction.
- **Removed the standalone `'GANA'` promotional keyword** — it was matching essentially at random
  inside long tracking-parameter strings in HTML email boilerplate (`z=1AvGy9akdi2...`), not real
  promotional text. `PARTICIPA Y GANA` (the full phrase) is unaffected.
- **New BANESCO `merchantPattern`**, rewritten from a real sample: `realizada en el ([^\n]+?) por` —
  the old pattern looked for "establecimiento/comercio/lugar", none of which the real template
  uses. Unconfirmed for an *approved* consumption alert yet (the only real sample was the declined
  one above, now excluded) — will need a real approved-charge sample to verify.
- **Every successful parse is now also logged**, with merchant + amount + a body snippet — the
  last real run saved 3 transactions with merchant `"<first name>"` (the customer's own name from the
  email greeting, not the actual store), confirming `LAFISE.merchantPattern` still needs a real
  sample to fix properly. This log line means the next run surfaces what's needed without asking
  for another manual copy-paste.

### v1.0.7 (Sept 23, 2026) — verbose rejection logging
- 13 real emails were found (confirming the v1.0.5 sender-domain fix works) but 0 parsed —
  meaning `extractAmount()`'s regexes, also written without ever seeing a real email, likely
  don't match the actual amount-line format your banks use.
- `parseEmailMessage()` now logs a body snippet whenever amount extraction fails, and
  `isPromotionalEmail()` / `isAccountNoticeEmail()` now log which specific keyword matched when
  they reject something. Run "Monitor Gmail Now" (or re-run the search) and check View > Logs —
  the exact reason for each of the 13 rejections is now visible without opening Gmail manually.

### v1.0.6 (Sept 23, 2026) — filter account/security notices, not just promos
- BDI's search is domain-wide (`from:bdi.com.do`) since we don't yet have a consumption-specific
  sender for it, which meant OTP codes, login alerts, and "beneficiario agregado" notices — real
  bank mail, just not transactions — could reach the parser. If one of those happened to mention
  a dollar figure in passing (an account limit, say), `extractAmount()` could have logged it as a
  phantom transaction.
- **New `ACCOUNT_NOTICE_KEYWORDS` / `isAccountNoticeEmail()`**, checked right after the
  promotional filter in `parseEmailMessage()`. Same mechanism as `PROMOTIONAL_KEYWORDS`, separate
  list/log line so you can tell from `View > Logs` which filter caught what.
- Once BDI's real "Aviso de Consumo" sender is confirmed (see the v1.0.5 note — hasn't shown up
  yet since no BDI card spend so far), pin `BDI.searchQuery` to that exact address the same way
  LAFISE/BANESCO are; this keyword filter stays on regardless as a second layer.

### v1.0.5 (Sept 23, 2026) — search by sender, not by keyword-in-body
- **Root cause of the Uber/Uber Eats contamination:** the search matched the bank's name
  anywhere in the email (subject **or body**), and Uber's own receipts mention the card's bank
  as the payment method (e.g. "charged to your LAFISE card"). A real `debugBankEmailSample()`
  run confirmed it: Uber Receipts (`noreply@uber.com`) showed up under both LAFISE and BHD,
  a LinkedIn invitation showed up under POPULAR, and even a self-sent email showed up under BDI
  — none of them from a bank.
- **Fix:** identification now happens by the message's actual sender domain
  (`BANK_PATTERNS[bank].fromDomain`), and the Gmail search itself is restricted with `from:`
  (`BANK_PATTERNS[bank].searchQuery`) instead of a bareword bank name. Confirmed senders so far:
  - **LAFISE** — `notificaciones@bancolafise.com`
  - **BANESCO** — `notificaciones@banesco.com.do` (pinned to this exact address, not the whole
    `banesco.com.do` domain — BANESCO's own marketing also comes from that domain, just a
    different address: `banescontigo@banesco.com.do`)
  - **BDI** — domain `bdi.com.do` (`bdinforma@bdi.com.do` confirmed for account-management mail;
    no consumption alert seen yet since BDI hasn't been used for card spend)
  - **BHD** — domain `bhd.com.do` minus the one confirmed marketing address
    (`info@bhd.com.do`), provisional until a real "Alerta de Consumo" email is seen
  - **POPULAR** — disabled (`searchQuery: null`). Every POPULAR hit in the diagnostic sample was
    a false positive; no real sender confirmed yet. Uncheck it in the Setup Wizard if you don't
    actually use it for spending — otherwise share a real sender address to re-enable it.
- This also fixes promotional bleed-through for banks whose marketing shares a domain with real
  alerts (BANESCO) — no more reliance on guessing every possible marketing phrase.

### v1.0.4 (Sept 23, 2026) — fixed the "0 saved" bug from v1.0.3
- **Root cause:** v1.0.3 narrowed the Gmail search to exact phrases ("aviso de consumo" /
  "notificación de consumo") that were a guess, not confirmed against real bank emails — and
  it still required `is:unread`, so any bank email you'd already opened (the normal case) was
  invisible to every search, regardless of keywords.
- **Fix:** search terms widened back to a high-recall list (`transacción OR pago OR consumo OR
  compra OR transferencia`); `is:unread` dropped entirely — duplicate protection is
  `isDuplicate()` in `04_sheetsWriter.gs` (matches by date + bank + amount), which already made
  the unread requirement unnecessary. The search is now bounded by a rolling date window
  (`daysBack`, default 60) instead, so it stays fast without depending on read/unread state.
  Promotional and $0.00 filtering still happen, but only in `parseEmailMessage()` (inspecting
  the full body) — more reliable than query-level phrase/negative-keyword matching, which is
  what caused the false-empty result.
- **New `debugBankEmailSample()`** — run it directly from the Apps Script editor (function
  dropdown → run → View > Logs) any time the monitor comes back empty. It searches by bank name
  only, no keyword filtering, and logs real subject lines so `BANK_PATTERNS` / `TYPE_KEYWORDS` /
  the amount regex can be tuned to match your actual emails instead of guessed phrasing.
- **"Monitor Gmail Now" alert now shows the funnel**: correos encontrados → transacciones
  parseadas → guardadas/duplicados/fallidas, so a future "0 saved" tells you which stage failed
  instead of just the final count.

### v1.0.3 (Sept 23, 2026) — Excel companion, categories rewrite, cleaner Gmail capture
- **New:** `Financial_Tracker_2026.xlsx` — a standalone workbook with filterable raw-transaction
  tabs per bank (`Raw_LAFISE`, `Raw_BANESCO`, `Raw_BHD`, `Raw_BDI`), a `Categorias` config tab,
  and a `Dashboard` that auto-totals spend by category and month via `SUMIFS`/`SUMPRODUCT`
  formulas (no hardcoded numbers — edit the raw tabs and the Dashboard recalculates). Useful on
  its own, or as a manual bridge until the Gmail monitor is fully dialed in.
- **Categories rewritten** (13, now matching the Excel exactly) — see the section below.
  `Cashback` was removed as a category; it's now a `Type` value instead.
- **New `Type` column** (`Transacción` / `Transferencia` / `Pago TC` / `Cashback`), appended as
  column K in the `Transactions` sheet and auto-detected by `detectTransactionType()` in
  `03_gmailMonitor.gs`. Category is only set for `Type = Transacción` — transfers and card
  payoffs stay uncategorized so they don't inflate spending totals.
- **BDI Digital added** to the Gmail bank patterns (was missing before — LAFISE, BANESCO, BHD,
  and POPULAR only).
- **Gmail search narrowed to consumption-only** notifications, with a promotional-keyword
  exclusion list (`isPromotionalEmail()`) and a $0.00 amount filter, so newsletters, offers, and
  zero-amount alerts never reach the sheet.
- If you deployed before this version: your `Transactions` sheet is missing the `Type` column —
  add a header `Type` in `K1` manually, or delete the sheet and let
  `initializeTransactionsSheet()` recreate it (you'll lose logged rows).

### v1.0.2 (Sept 21, 2026) — README fix
- Fixed: Quick Start told you to create a standalone project at script.google.com and paste a "spreadsheet ID" that never existed in the code. The script uses `getActiveSpreadsheet()` and an `onOpen()` menu trigger, both of which require the script to be **container-bound** (created via the Sheet's `Extensions > Apps Script` menu). Rewrote Steps 1–2 accordingly — no ID to find or paste anywhere.

### v1.0.1 (Sept 21, 2026) — Bug fixes
- Fixed: `extractMerchant()` was defined twice (01_main.gs and 03_gmailMonitor.gs) with incompatible signatures, causing a crash on every "Monitor Gmail Now" run.
- Fixed: `categorizeTransaction()` and `saveTransaction()` were also duplicated across files; removed the outdated copies from `01_main.gs` so there's a single source of truth for each.
- Fixed: `runGmailMonitor()` now calls the real pipeline (`searchTransactionEmails` → `extractTransactionsFromThreads` → `saveTransactions` → `markEmailsAsProcessed`) instead of its own out-of-date inline logic.
- Fixed: `Transactions` sheet header now has 10 columns (added `IsCredit`, `IsCashback`) to match what `saveTransaction()` actually writes — previously the header had 8 columns while 10 values were being appended.
- Fixed: `CustomRules` sheet is now created automatically during Setup Wizard, so `addCustomCategoryRule()`/`getUserCustomRules()` won't throw when first used.
- Hardened: all sheet lookups now use `getOrCreateSheet()` instead of `getSheetByName()` so a missing sheet no longer causes an unhandled error.
