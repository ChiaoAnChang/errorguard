# Synthetic test data

`sample_data.csv` is entirely fabricated (fictional names, made-up numbers) — no real person, organization, or dataset. It's small on purpose, built to exercise every v1 rule at once when pasted into a fresh Excel sheet:

- **Duplicate row:** row 2 and row 6 (`Alice Johnson`) are identical.
- **Mixed date formats:** row 3 (`01/16/2024`) vs. the rest of the column (`YYYY-MM-DD`).
- **Text in a numeric column:** row 7 (`N/A`) in an otherwise-numeric `Amount` column.
- **Inconsistent casing:** row 8 (`ACTIVE`) vs. the rest of `Status`; row 9 (`grace kim`) vs. the rest of `Name`.
- **Leading/trailing whitespace:** row 10 (`" 75.00"`) and row 12 (`" 2024-01-24 "`).
- **Statistical outlier:** row 12's `500000.00` against an otherwise ~$45–$210 `Amount` column.

## Manual sideload test (spec execution plan, phase 5)

1. `npm start` to sideload ErrorGuard into Excel Desktop (or open Excel on the web and sideload manually — see the root [README.md](../README.md)).
2. Open a blank workbook, paste this CSV's contents into `Sheet1` starting at `A1` (Excel will parse the header row automatically).
3. Open the ErrorGuard task pane and confirm all six issues above are listed with a sensible reason, and the corresponding cells are highlighted.
4. Dismiss one flag, confirm it grays out and stops being highlighted; edit that same cell and confirm the flag reappears (a new flag id, since the cell changed).
5. Repeat in Excel on the web.
