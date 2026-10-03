# Tests — Financial Tracker

```
node --test tests/*.test.js          # Node 18+; no dependencies
npm test                             # same thing, via package.json
GS_ROOT=/other/copy node --test tests/*.test.js   # run against another copy of the .gs files
```

| File | Covers |
|---|---|
| `harness.js` | Loads the four `.gs` files into ONE shared scope (like Apps Script) with small stubs for `Logger`, `Session`, `Utilities`. |
| `sheets_mock.js` | Strict mock of `SpreadsheetApp`, `GmailApp`, `LockService`: unknown API methods throw, writes outside the sheet's size throw, `getUi()` throws (as in a trigger). |
| `categorizer.test.js` | Keyword engine, categories, Custom Rules sanitizing, custom-only categories. |
| `parsing.test.js` | Type detection, every fixture end-to-end through `parseEmailMessage()`, declined rows, balances, IsCredit, garbled text, filters. |
| `v124.test.js` | DR payroll (published RD$50,000 example, caps, top bracket, 2027 flag), setup validation, wizard with no personal data, save → Configuration + triggers, ISR row, generic default cards, column auto-fit, daily summary (numbers, email, delivery on/off). |
| `v126.test.js` | Day-to-day comparison without bills, redesigned daily email (blocks, preheader, links, data health), monthly summary (totals, comparisons, subscriptions, missed cashback, year to date, calendar), monthly trigger and on/off, run status recorded. |
| `v127.test.js` | Other income (Dashboard, emails), card catalogue and configured cards, Setup Wizard save flow (toasts, summary, errors after the window closes). |
| `v128.test.js` | Category colours (distinct, WCAG AA contrast), data-sheet rules and their priority, Bank Transfers Category column and migration, Custom Rules dropdown, Categories colours. |
| `v129.test.js` | Investments: HAPI email parsing (incl. HTML-only and inconsistent orders), holdings maths, a full run from emails and bank deposits to Holdings, run summary. |
| `v132.test.js` | Investments reporting: XIRR, Modified Dietz returns per account and in total, history rows (one set per day), daily and monthly briefs, email sections, Holdings chart not stacking. |
| `v133.test.js` | No progress toast left open (success and error paths), fund/pension balance dialog: validation, saving, listing, Holdings. |
| `v134.test.js` | Código Cash merchants (parse and repair), Holdings dates as real dates, the Dashboard look of Holdings and the other investment tabs. |
| `v135.test.js` | Unrecognized sheet: listing, one row per email, Status kept, resolved emails removed, broker failures, data-health line; balances updated not duplicated. |
| `v136.test.js` | Time budget: saved emails not read again, reading stops at the deadline, a stopped run finishes its steps and the next run completes it; broker emails likewise. |
| `v137.test.js` | Investment changes account by account (added accounts aren't gain), status colours vs chips (ΔE ≥ 25), tab order and the settings-tab toggle, link noise in Unrecognized. |
| `v138.test.js` | HAPI limit orders (Cost with the fee), the limit-order email leaving Unrecognized, tab order keeping hidden tabs hidden, Recategorize reordering tabs. |
| `v139.test.js` | Hand-set categories kept (Transactions and Bank Transfers, before and after v1.1.39), rules vs your choices, HAPI's return from an earlier balance, history day bands. |
| `v140.test.js` | Missing-deposit check (purchases vs deposits since the start), its warning in Performance, deposits from the dialog. |
| `v141.test.js` | Every dialog's page script compiles as the browser gets it, and its buttons and options call functions that exist. |
| `v142.test.js` | Deposit notices: pairing with recorded deposits (window, one each, same account), no effect on money, Unrecognized rows that clear when the deposit is added, the run path. |
| `v143.test.js` | Read log (read-only emails not read again; re-read after an update; failures retried), stopped runs defer heavy steps, the 5-minute brake, undated ledger rows, the deposit date. |
| `v144.test.js` | Marking only pending threads, the marking time budget, per-phase timings, start-day sale proceeds as start cash. |
| `bundle.test.js` | `dist/FinancialTracker.gs` is what `src/` builds; every file in load order; compiles as one script. (`npm run test:bundle` runs the whole suite against that file.) |
| `v147.test.js` | 📘 Start here: a new sheet, a sheet with real problems, unscheduled runs, the sidebar's script and allowed actions, allowed tabs, menu and wizard pointer. |
| `v148.test.js` | Pasting a broker's portfolio screen: parsing (plain and linked), Total assets check, cash, duplicates, saving and same-day replacement, entry points. |
| `v149.test.js` | Deleted rows re-imported, Reset System re-importing everything and keeping investments, the Spanish guide in sync with the menu, the monthly note before the investment history. |
| `v150.test.js` | Every "📊 Tracker › …" the tracker shows names a menu item that exists. |
| `v151.test.js` | Incoming transfers: LAFISE emails (third party, own account, Custom Rule), Banesco statements in three text layouts and their totals, a run via Drive, netting in the summary, recategorizing, no Drive API. |
| `v152.test.js` | The PDF reader: inflate vs zlib, a FOP-style statement PDF to lines (glyph widths, TJ spacing, Type0 title), PDF values, unsupported files; LAFISE announcements filtered. |
| `v153.test.js` | Incoming Transfers gets what Bank Transfers gets (styles, fitted columns, date order) and is deleted by Reset System. |
| `v154.test.js` | Categories typed in the transfer sheets reach Transactions at once; rows typed in Incoming Transfers are saved (once, category kept) or listed in Unrecognized. |
| `v151.test.js` (v1.1.55) | Every statement credit; a statement imported by an earlier version read again without duplicates. Fixtures: `fixtures/statements/make_statement_text.py` → text, `make_synthetic_pdf.py` → PDF. |
| `v156.test.js` | Waiting for live prices (loads, never loads, time limit), unpriced positions, the Day change column, session moves, the email note, and a refresh recording the loaded price. |
| `v157.test.js` | Transfers sent typed in Bank Transfers: saved once, category kept or set by a Custom Rule; rows of pre-v1.1.39 sheets never duplicated. |
| `v158.test.js` | Typed transfers keep their date (real dates and text), no day shift on yyyy-mm-dd text, rows saved without a date repaired instead of duplicated, the check after saving. |
| `v159.test.js` | Other incomes and deductions as lines (net income, validation, saving, Dashboard rows), an older single other income migrated (Dashboard included), protected PDFs named. |
| `v162.test.js` | Gmail's bold markers (`*text*`) on real-structure fixtures and on every newer extractor; QIK purchases, LAFISE's second template (DOP and USD), the first template's currency; BDI card payment and its tax row; emails that aren't a movement (BDI "Completada", QIK code created, payroll, loan reminder, validation code); BANESCO transfers received and their repair; the Dashboard's arithmetic computed over every category, currency and type; categories typed by hand. |
| `v161.test.js` | LAFISE card payments and its second card template (both senders searched), QIK Código CASH, transfers received (BDI, Scotiabank, Banreservas), Banreservas transfers and TuEfectivo withdrawals, strict banks never guessed, welcome emails filtered, unticked banks named in the summary, counts by bank. |
| `v160.test.js` | BDI (purchases, interbank transfers sent), Scotiabank and QIK from real-sample structure; a security footer not filtering a purchase; one list of banks; the new catalogue products. |
| `pipeline.test.js` | Date windows, dedup, recategorize rules, run summary. |
| `integration.test.js` | Full runs against the mock: search bounds, month-spanning threads, raw-sheet rebuild, Notes, run lock, 2,000+ rows; Dashboard formula lint (syntax, real functions, existing named ranges), row wiring, migration from the v1.1.20 layout and preservation of every input. |

## Adding a real email as a fixture

1. Save the email's plain text as `fixtures/<bank>_<kind>_<short-name>.txt`.
2. **Anonymize before saving**: your name, card/account digits, confirmation numbers.
   Keep the layout (line breaks, labels, currency format) exactly as it is — that's what the test is for.
3. Add an entry to `fixtures/manifest.json` with subject, sender, date, `"synthetic": false` and the
   expected items (merchant, amount, currency, type, category). An email that must produce nothing
   gets `"expected": []`.

## Manual checks (the mock can't cover these)

- [ ] View → Executions shows `1.1.19` on a fresh run.
- [ ] 📅 Monitor by Date Range: window closes after ~1.5 s, toasts appear bottom-right, summary alert at the end.
- [ ] Clicking Run twice quickly → second run shows "Another run is already in progress".
- [ ] Dashboard looks right: KPI cards, bars, heat map, chart; Month = "Current month" shows this month.
- [ ] Editing a rate (USD/EUR/COP) stamps today's date in "Rates updated".
- [ ] "Google ref." under the rates shows values or "n/a" (reference only).
- [ ] Next 6 AM trigger run appears in Executions without errors.
- [ ] Setup Wizard with DOP + automatic shows the ARS/AFP/ISR preview; saving creates the summary trigger (Triggers page).
- [ ] 🗓️ Send Monthly Summary Now → the email for last month arrives and looks right on phone and desktop.
- [ ] 📬 Send Daily Summary Now → the email arrives, looks right on phone and desktop, and its button opens the Dashboard.
- [ ] Gmail: processed threads carry the **Procesado** label; a thread reported as "Could not parse" stays unread.
