# Bank ↔ Mygate Reconciliation

A single-file [Streamlit](https://streamlit.io) app that reconciles the
quarterly **ICICI bank statements** against the **Mygate books ledger** for a
residential society, and highlights matched transactions **green** in both.

## What it does

1. Upload the **Mygate ledger** export (`.xls` / `.xlsx`).
2. Upload one or more **bank statements** (`.xls` / `.xlsx`).
3. The app matches every bank transaction to the ledger and returns a **ZIP**
   of the same sheets with matched rows highlighted green, plus a summary.

Files are processed **in memory** — nothing is stored on the server, and your
originals are never modified.

## Matching logic

Key = **date + amount + direction**, matched **one-to-one**:

| Bank side | Direction | Mygate side |
|-----------|-----------|-------------|
| Deposit (money in)     | `in`  | Debit to the bank ledger  |
| Withdrawal (money out) | `out` | Credit to the bank ledger |

Descriptions are deliberately ignored (the two systems word them
differently). Recurring amounts — e.g. a maintenance charge paid by many
flats — are handled by consuming candidate pairs in order of **increasing
date gap**, so each bank line is paired with its closest available book entry.
A wide fallback window (default **120 days**) still lets a payment booked a
quarter later find its match.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py        # → http://localhost:8501
```

## Deploy to Streamlit Community Cloud

1. Push this repo to GitHub (already done).
2. Go to <https://share.streamlit.io> → **Create app** → **Deploy from GitHub**.
3. Repository `nitish31jain/bank_and_mygate_recon`, branch `main`,
   **main file `app.py`**.
4. Deploy — it installs `requirements.txt` automatically.

## Files

| File | Purpose |
|------|---------|
| `app.py` | The entire app — reconciliation logic + Streamlit UI |
| `requirements.txt` | `streamlit`, `openpyxl` |

> Actual bank statements and the Mygate export contain private financial data
> and are **not** committed (see `.gitignore`). Users upload their own files
> at runtime.
