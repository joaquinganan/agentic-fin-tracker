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
