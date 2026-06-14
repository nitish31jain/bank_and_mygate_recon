# Bank ↔ Mygate Reconciliation (FY 2025-26)

Reconciles the quarterly **ICICI bank statements** against the **Mygate books
ledger** for Vaswani Brentwood, FY 2025-26, and highlights matched rows
**green** in both the bank files and the Mygate file (originals are
overwritten in place).

## Matching logic

The key is **date + amount + direction**, matched **one-to-one**:

| Bank side | Direction | Mygate side |
|-----------|-----------|-------------|
| Deposit (money in)     | `in`  | Debit to the bank ledger  |
| Withdrawal (money out) | `out` | Credit to the bank ledger |

Transaction **descriptions are deliberately ignored** — the bank and Mygate
word the same transaction completely differently, so amount + date +
direction is the only reliable key.

### Nearest-date-first, one-to-one

Recurring amounts (e.g. the maintenance charge `20,559` paid by many flats)
appear many times. The script builds every candidate `(bank, mygate)` pair
that shares the same amount + direction within a date window, then consumes
them in **order of increasing date gap**. This guarantees each bank line is
paired with its *closest available* book entry rather than an arbitrary
same-amount one, while a wide window (`MAXWIN`, default **120 days**) still
lets a payment booked a quarter later find its match.

## Usage

```bash
pip install openpyxl
python reconcile.py
```

The data files are expected in the working directory:

- `Mygate Transactions FY 2025-26.xls` — Mygate ledger export
  (note: this is actually `.xlsx` content saved with a `.xls` name; the
  script handles that)
- `Apr_Jun_2025.xlsx`, `Jul_Sep_2025.xlsx`, `Oct_Dec_2025.xlsx`,
  `Jan_Mar_2026.xlsx` — quarterly bank statements

> ⚠️ Close the Excel files before running — Windows locks open workbooks and
> the save will fail with a permission error.

After running, **matched rows are green**; **un-highlighted rows are the
exceptions** to investigate.

## Output

The script prints a summary: matches bucketed by date gap (same-day / 1–5 /
6–30 / 31–90 / 91–120 days) plus matched / unmatched counts per file and
totals for both sides.

## Note on the data files

The actual bank statements and Mygate export contain private financial data
and are **not** committed (see `.gitignore`). Only the reconciliation code is
in this repository.
