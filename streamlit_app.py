"""
Streamlit front-end for bank <-> Mygate reconciliation.

Reuses the in-memory engine in reconcile_core.py. Upload the Mygate ledger +
one or more bank statements, reconcile, and download a ZIP of the same sheets
with matched rows highlighted green, plus a summary.

Local run:        streamlit run streamlit_app.py
Deploy:           push to GitHub, then https://share.streamlit.io -> New app
                  -> pick repo/branch, main file = streamlit_app.py
"""

import io
import os
import zipfile

import streamlit as st

from reconcile_core import reconcile, DEFAULT_MAXWIN

st.set_page_config(page_title="Bank ↔ Mygate Reconciliation", page_icon="✅",
                   layout="centered")

st.title("Bank ↔ Mygate Reconciliation")
st.caption(
    "Upload the Mygate ledger and your bank statements. Matched transactions "
    "are highlighted green in both, and you download the result as a ZIP. "
    "Files are processed in memory — nothing is stored on the server."
)

with st.form("recon"):
    mygate_file = st.file_uploader(
        "Mygate ledger file (single .xls / .xlsx)",
        type=["xls", "xlsx"], accept_multiple_files=False)
    bank_files = st.file_uploader(
        "Bank statement file(s) (.xls / .xlsx) — select one or more",
        type=["xls", "xlsx"], accept_multiple_files=True)
    maxwin = st.number_input(
        "Date-gap window for fallback matches (days)",
        min_value=0, max_value=366, value=DEFAULT_MAXWIN, step=5,
        help="Recurring amounts (e.g. maintenance charges) may be booked weeks "
             "apart from the bank date. Matches are still made nearest-date-first.")
    submitted = st.form_submit_button("Reconcile & build ZIP")

if submitted:
    if not mygate_file:
        st.error("Please choose the Mygate ledger file.")
        st.stop()
    if not bank_files:
        st.error("Please choose at least one bank statement file.")
        st.stop()

    mygate = (mygate_file.name, mygate_file.getvalue())
    banks = [(f.name, f.getvalue()) for f in bank_files]

    try:
        with st.spinner("Reconciling…"):
            output_files, s = reconcile(mygate, banks, int(maxwin))
    except Exception as exc:
        st.error(f"Reconciliation failed: {exc}")
        st.stop()

    rate = (s["bank_matched"] / s["bank_total"] * 100) if s["bank_total"] else 0
    st.success(f"Done — {rate:.1f}% of bank lines matched "
               f"({s['bank_matched']} of {s['bank_total']}).")

    c1, c2, c3 = st.columns(3)
    c1.metric("Bank matched", f"{s['bank_matched']}", f"-{s['bank_unmatched']} unmatched",
              delta_color="inverse")
    c2.metric("Mygate matched", f"{s['mygate_matched']}", f"-{s['mygate_unmatched']} unmatched",
              delta_color="inverse")
    c3.metric("Match rate", f"{rate:.1f}%")

    b = s["buckets"]
    st.subheader(f"Matches by date gap (window {s['maxwin']} days)")
    st.table({
        "Gap": ["same day", "1–5 days", "6–30 days", "31–90 days", "91+ days"],
        "Matches": [b["same_day"], b["1_5"], b["6_30"], b["31_90"], b["91_plus"]],
    })

    st.subheader("Per file")
    rows = [{"File": pf["name"], "Transactions": pf["total"],
             "Matched": pf["matched"], "Unmatched": pf["unmatched"]}
            for pf in s["per_file"]]
    rows.append({"File": "BANK TOTAL", "Transactions": s["bank_total"],
                 "Matched": s["bank_matched"], "Unmatched": s["bank_unmatched"]})
    rows.append({"File": s["mygate_name"], "Transactions": s["mygate_total"],
                 "Matched": s["mygate_matched"], "Unmatched": s["mygate_unmatched"]})
    st.dataframe(rows, use_container_width=True, hide_index=True)

    summary_text = "\n".join([
        "BANK <-> MYGATE RECONCILIATION SUMMARY",
        "=" * 50,
        f"Date-gap window: {s['maxwin']} days",
        f"Bank   : matched {s['bank_matched']} / {s['bank_total']} "
        f"(unmatched {s['bank_unmatched']})",
        f"Mygate : matched {s['mygate_matched']} / {s['mygate_total']} "
        f"(unmatched {s['mygate_unmatched']})",
        "",
        "Matches by date gap:",
        f"  same day : {b['same_day']}",
        f"  1-5 days : {b['1_5']}",
        f"  6-30 days: {b['6_30']}",
        f"  31-90    : {b['31_90']}",
        f"  91+      : {b['91_plus']}",
    ])

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in output_files:
            zf.writestr(os.path.basename(name), data)
        zf.writestr("reconciliation_summary.txt", summary_text)
    zip_buf.seek(0)

    st.download_button(
        "⬇ Download reconciled sheets (ZIP)",
        data=zip_buf.getvalue(),
        file_name="reconciled_sheets.zip",
        mime="application/zip",
        type="primary",
    )
    st.caption("Green-highlighted rows matched. Un-highlighted rows are the "
               "exceptions to review.")
