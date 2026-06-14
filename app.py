"""
Flask web app for bank <-> Mygate reconciliation.

Upload the Mygate ledger export + one or more bank statements, the app
reconciles them and returns a ZIP of the same files with matched rows
highlighted green, plus a summary.txt.

Run:  python app.py    then open http://127.0.0.1:5000
"""

import io
import os
import zipfile
import secrets

from flask import (Flask, request, render_template, send_file,
                   redirect, url_for, flash, abort)

from reconcile_core import reconcile, DEFAULT_MAXWIN

app = Flask(__name__)
app.secret_key = secrets.token_hex(16)
app.config["MAX_CONTENT_LENGTH"] = 64 * 1024 * 1024  # 64 MB total upload cap

# in-memory store of generated results: token -> dict(zip_bytes, summary)
RESULTS = {}
MAX_RESULTS = 20


def _remember(zip_bytes, summary):
    token = secrets.token_urlsafe(12)
    RESULTS[token] = dict(zip=zip_bytes, summary=summary)
    # keep only the most recent few
    while len(RESULTS) > MAX_RESULTS:
        RESULTS.pop(next(iter(RESULTS)))
    return token


def _build_summary_text(s):
    b = s["buckets"]
    lines = [
        "BANK <-> MYGATE RECONCILIATION SUMMARY",
        "=" * 50,
        f"Date-gap window for fallback matches: {s['maxwin']} days",
        "",
        "Matches by date gap:",
        f"  same day    : {b['same_day']}",
        f"  1-5 days    : {b['1_5']}",
        f"  6-30 days   : {b['6_30']}",
        f"  31-90 days  : {b['31_90']}",
        f"  91+ days    : {b['91_plus']}",
        "",
        "Per bank file:",
    ]
    for pf in s["per_file"]:
        lines.append(f"  {pf['name']:<28} txns {pf['total']:>5} | "
                     f"matched {pf['matched']:>5} | unmatched {pf['unmatched']:>4}")
    lines += [
        "-" * 50,
        f"  {'BANK TOTAL':<28} txns {s['bank_total']:>5} | "
        f"matched {s['bank_matched']:>5} | unmatched {s['bank_unmatched']:>4}",
        f"  {s['mygate_name']:<28} entr {s['mygate_total']:>5} | "
        f"matched {s['mygate_matched']:>5} | unmatched {s['mygate_unmatched']:>4}",
        "",
        "Green-highlighted rows matched. Un-highlighted rows are exceptions to review.",
    ]
    return "\n".join(lines)


@app.route("/")
def index():
    return render_template("index.html", default_maxwin=DEFAULT_MAXWIN)


@app.route("/reconcile", methods=["POST"])
def do_reconcile():
    mygate_file = request.files.get("mygate")
    bank_files = [f for f in request.files.getlist("banks") if f and f.filename]

    if not mygate_file or not mygate_file.filename:
        flash("Please choose the Mygate ledger file.")
        return redirect(url_for("index"))
    if not bank_files:
        flash("Please choose at least one bank statement file.")
        return redirect(url_for("index"))

    try:
        maxwin = int(request.form.get("maxwin", DEFAULT_MAXWIN))
    except ValueError:
        maxwin = DEFAULT_MAXWIN
    maxwin = max(0, min(maxwin, 366))

    mygate = (mygate_file.filename, mygate_file.read())
    banks = [(f.filename, f.read()) for f in bank_files]

    try:
        output_files, summary = reconcile(mygate, banks, maxwin)
    except Exception as exc:  # surface a friendly message
        flash(f"Reconciliation failed: {exc}")
        return redirect(url_for("index"))

    summary_text = _build_summary_text(summary)

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in output_files:
            zf.writestr(os.path.basename(name), data)
        zf.writestr("reconciliation_summary.txt", summary_text)
    zip_buf.seek(0)

    token = _remember(zip_buf.getvalue(), summary)
    return render_template("result.html", summary=summary, token=token,
                           summary_text=summary_text)


@app.route("/download/<token>")
def download(token):
    item = RESULTS.get(token)
    if not item:
        abort(404, "Result expired or not found. Please re-run the reconciliation.")
    return send_file(io.BytesIO(item["zip"]), mimetype="application/zip",
                     as_attachment=True, download_name="reconciled_sheets.zip")


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
