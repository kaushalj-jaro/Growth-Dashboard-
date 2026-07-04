"""
app.py
------
The upload portal. Run this, open it in a browser, upload your Raw CSV
and Enrolled CSV (and Targets .xlsx the first time / whenever targets
change) -- it runs process_reports.py end to end and writes every
dashboard straight into Neon Postgres. Nothing else to run in between.

Start it with:
    uvicorn app:app --reload

Then open:
    http://127.0.0.1:8000
"""

import os
import shutil
import traceback
from datetime import datetime

from fastapi import FastAPI, File, UploadFile
from fastapi.responses import HTMLResponse
from sqlalchemy import text

from process_reports import run_pipeline
from database import engine

app = FastAPI(title="Growth Funnel Report Portal")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

# The targets workbook rarely changes -- if the user doesn't upload a new
# one this run, we fall back to whichever one was last saved here.
LAST_TARGETS_PATH = os.path.join(UPLOAD_DIR, "_last_targets.xlsx")


PAGE_HEAD = """
<html>
<head>
<title>Growth Funnel Report Portal</title>
<style>
  body { font-family: -apple-system, Arial, sans-serif; max-width: 640px; margin: 60px auto; color: #1a1a1a; }
  h1 { font-size: 22px; }
  label { display: block; margin-top: 18px; font-weight: 600; font-size: 14px; }
  input[type=file] { margin-top: 6px; }
  button { margin-top: 26px; padding: 10px 22px; font-size: 15px; border: none;
           border-radius: 6px; background: #d97757; color: white; cursor: pointer; }
  button:hover { background: #c96a4d; }
  .note { color: #666; font-size: 13px; margin-top: 4px; }
  .success { background: #eafaf0; border: 1px solid #b7e4c7; padding: 16px; border-radius: 8px; }
  .error { background: #fdecea; border: 1px solid #f5b7b1; padding: 16px; border-radius: 8px; white-space: pre-wrap; font-family: monospace; font-size: 13px; }
  table { border-collapse: collapse; margin-top: 12px; width: 100%; }
  td, th { text-align: left; padding: 4px 10px; border-bottom: 1px solid #eee; font-size: 14px; }
  a { color: #d97757; }
</style>
</head>
<body>
"""

PAGE_TAIL = "</body></html>"

FORM_HTML = f"""
{PAGE_HEAD}
<h1>Growth Funnel Report Portal</h1>
<p>Upload today's files. This runs your full processing pipeline and writes every dashboard straight into Neon Postgres.</p>
<form action="/upload" method="post" enctype="multipart/form-data">
  <label>Raw Data CSV</label>
  <input type="file" name="raw_file" accept=".csv" required>

  <label>Enrolled Data CSV</label>
  <input type="file" name="enrolled_file" accept=".csv" required>

  <label>Targets Workbook (.xlsx)</label>
  <input type="file" name="targets_file" accept=".xlsx">
  <div class="note">Optional -- only upload this when your targets change. Otherwise the last one you uploaded will be reused automatically.</div>

  <button type="submit">Process &amp; Save to Postgres</button>
</form>
<p style="margin-top:30px;"><a href="/history">View recent runs</a></p>
{PAGE_TAIL}
"""


@app.get("/", response_class=HTMLResponse)
def form():
    return FORM_HTML


@app.post("/upload", response_class=HTMLResponse)
async def upload(
    raw_file: UploadFile = File(...),
    enrolled_file: UploadFile = File(...),
    targets_file: UploadFile = None,
):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    raw_path = os.path.join(UPLOAD_DIR, f"{timestamp}_{raw_file.filename}")
    enrolled_path = os.path.join(UPLOAD_DIR, f"{timestamp}_{enrolled_file.filename}")

    with open(raw_path, "wb") as f:
        shutil.copyfileobj(raw_file.file, f)
    with open(enrolled_path, "wb") as f:
        shutil.copyfileobj(enrolled_file.file, f)

    # Use the freshly uploaded targets file if given, otherwise reuse the last one
    if targets_file is not None and targets_file.filename:
        with open(LAST_TARGETS_PATH, "wb") as f:
            shutil.copyfileobj(targets_file.file, f)

    if not os.path.exists(LAST_TARGETS_PATH):
        return f"""
        {PAGE_HEAD}
        <div class="error">No targets workbook found. Upload a Targets .xlsx at least once before running the pipeline.</div>
        <p><a href="/">&larr; Back</a></p>
        {PAGE_TAIL}
        """

    try:
        result = run_pipeline(raw_path, enrolled_path, LAST_TARGETS_PATH)
    except Exception:
        error_text = traceback.format_exc()
        return f"""
        {PAGE_HEAD}
        <h1>❌ Pipeline failed</h1>
        <div class="error">{error_text}</div>
        <p><a href="/">&larr; Back</a></p>
        {PAGE_TAIL}
        """

    rows_html = "".join(
        f"<tr><td>{name}</td><td>{count}</td></tr>"
        for name, count in result["sheets"].items()
    )

    return f"""
    {PAGE_HEAD}
    <h1>✅ Pipeline complete</h1>
    <div class="success">
      Cycle: {result['cycle_month']} {result['cycle_year']}<br>
      Upload ID: {result['upload_id']}
    </div>
    <table>
      <tr><th>Table</th><th>Rows</th></tr>
      {rows_html}
    </table>
    <p style="margin-top:20px;"><a href="/">&larr; Run another upload</a></p>
    {PAGE_TAIL}
    """


@app.get("/history", response_class=HTMLResponse)
def history():
    with engine.connect() as conn:
        rows = conn.execute(text("""
            SELECT upload_id, run_timestamp, cycle_month, cycle_year, raw_file, enrolled_file
            FROM upload_history
            ORDER BY run_timestamp DESC
            LIMIT 25
        """)).fetchall()

    rows_html = "".join(
        f"<tr><td>{r.run_timestamp}</td><td>{r.cycle_month} {r.cycle_year}</td>"
        f"<td>{os.path.basename(r.raw_file or '')}</td><td>{os.path.basename(r.enrolled_file or '')}</td></tr>"
        for r in rows
    )

    return f"""
    {PAGE_HEAD}
    <h1>Recent runs</h1>
    <table>
      <tr><th>Run time</th><th>Cycle</th><th>Raw file</th><th>Enrolled file</th></tr>
      {rows_html}
    </table>
    <p style="margin-top:20px;"><a href="/">&larr; Back to upload</a></p>
    {PAGE_TAIL}
    """
