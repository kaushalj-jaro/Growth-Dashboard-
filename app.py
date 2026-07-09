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

import json
import os
import shutil
import traceback
from datetime import datetime

import pandas as pd
from fastapi import FastAPI, File, UploadFile
from fastapi.responses import HTMLResponse, Response, JSONResponse
from sqlalchemy import text

from process_reports import run_pipeline
from daily_tracker import run_daily_tracker, backfill_daily_history, TABLE_MAP as DAILY_TABLE_MAP
from database import engine
from db_writer import SHEET_TABLE_MAP

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

  <label style="display:flex; align-items:center; gap:8px; font-weight:400; margin-top:18px;">
    <input type="checkbox" name="backfill_daily" value="yes" style="margin:0;">
    Backfill Daily Tracker history from the 1st of the month
  </label>
  <div class="note">
    Reconstructs one row per day using each lead's own Created On date. Delivered/Admissions
    counts are accurate historically; Workable/Prospect/Fresh/Junk for past days reflect each
    lead's CURRENT stage, not its actual stage back then (your CSV only stores one live
    snapshot, not a stage history) -- those days are tagged so they're never confused with a
    real day-by-day capture. Leave unchecked for a normal single-day update.
  </div>

  <button type="submit">Process &amp; Save to Postgres</button>
</form>
<p style="margin-top:30px;"><a href="/history">View recent runs</a> &nbsp;|&nbsp; <a href="/api/dashboards">API: list tables</a></p>
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
    backfill_daily: str = None,
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

    # Daily tracker runs separately from the monthly pipeline above -- if
    # this fails, we still show the monthly result as a success and just
    # surface the daily tracker error alongside it, instead of losing both.
    daily_result = None
    daily_error = None
    try:
        if backfill_daily == "yes":
            daily_result = backfill_daily_history(raw_path, enrolled_path)
        else:
            daily_result = run_daily_tracker(raw_path, enrolled_path)
    except Exception:
        daily_error = traceback.format_exc()

    rows_html = "".join(
        f"<tr><td>{name}</td><td>{count}</td></tr>"
        for name, count in result["sheets"].items()
    )

    daily_html = ""
    if daily_result and daily_result.get("backfilled"):
        start, end = daily_result["backfilled_range"]
        daily_html = f"""
        <div class="success" style="margin-top:16px;">
          Backfilled {daily_result['backfilled_days']} day(s): {start} → {end}<br>
          History now spans {daily_result['days_in_history']} day(s) total.<br>
          <span style="opacity:0.8;">Days before {daily_result['snapshot_date']} are reconstructed (accurate
          Delivered/Admissions, approximate Workable/Prospect/Fresh/Junk) -- tagged _backfilled in daily_snapshots.</span>
        </div>
        """
    elif daily_result:
        daily_html = f"""
        <div class="success" style="margin-top:16px;">
          Daily tracker snapshot: {daily_result['snapshot_date']}
          (raw as of {daily_result['raw_snapshot_date']}, enrolled as of {daily_result['enrolled_snapshot_date']})<br>
          History now spans {daily_result['days_in_history']} day(s).
        </div>
        """
    elif daily_error:
        daily_html = f"""
        <div class="error" style="margin-top:16px;">
          Daily tracker step failed (monthly dashboards above still saved fine):
          {daily_error}
        </div>
        """

    return f"""
    {PAGE_HEAD}
    <h1>✅ Pipeline complete</h1>
    <div class="success">
      Cycle: {result['cycle_month']} {result['cycle_year']}<br>
      Upload ID: {result['upload_id']}
    </div>
    {daily_html}
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


# ============================================================
# JSON API — for Google Apps Script / any external dashboard to pull
# the latest data. No auth for now (per your call) -- anyone with the
# URL can read this. Add an API key later if this ever leaves your
# own network.
# ============================================================

# Every table this endpoint is allowed to read. Deliberately an
# allowlist, not "whatever table name you pass in" -- avoids letting
# a URL query an arbitrary table in your database.
ALLOWED_TABLES = set(SHEET_TABLE_MAP.values()) | set(DAILY_TABLE_MAP.values()) | {"subsource_dashboard", "upload_history"}


@app.get("/api/dashboards")
def list_dashboards():
    """GET /api/dashboards -- lists every table name you can query below."""
    return JSONResponse(sorted(ALLOWED_TABLES))


@app.get("/api/dashboard/{table_name}")
def get_dashboard(table_name: str):
    """
    GET /api/dashboard/overall_dashboard
    Returns that table's current contents as JSON. Since every monthly
    table is fully overwritten on each pipeline run (if_exists="replace"),
    this always reflects the latest upload -- no extra filtering needed.

    From Apps Script:
        UrlFetchApp.fetch("http://YOUR_SERVER:8000/api/dashboard/overall_dashboard")
    """
    if table_name not in ALLOWED_TABLES:
        return JSONResponse(
            {"error": f"'{table_name}' isn't a recognized table.", "available": sorted(ALLOWED_TABLES)},
            status_code=404,
        )

    try:
        df = pd.read_sql(f'SELECT * FROM "{table_name}"', engine)
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)

    # to_json (not a plain dict) so Timestamps/NaN/numpy types serialize cleanly
    return Response(content=df.to_json(orient="records", date_format="iso"), media_type="application/json")


# ============================================================
# /dashboard/{month} — the endpoint your Apps Script Code.gs is
# ALREADY calling (as "https://.../dashboard/july"). Code.gs needs
# no changes; this is the piece that was missing on this side.
#
# Shape must match sheet.getDataRange().getValues(): each tab is an
# array of rows, the first row being headers -- so Code.gs's
# monthlyPayload["July"] = julyData.data behaves identically whether
# a tab came from a real Google Sheet or from here.
# ============================================================

# Exact tab names your Code.gs's EXPECTED_TABS list asks for.
EXPECTED_TABS = [
    "Enhanced Branch Report",
    "Enhanced University Report",
    "Overall Dashboard",
    "PR Dashboard",
    "Whatsapp Dashboard",
    "AI Meeting Dashboard",
    "Website Dashboard",
    "Live Chat Dashboard",
    "RDMPL Dashboard",
    "Website Campaign Dashboard",
    "Subsource Dashboard",
    "Product Dashboard",
]

# Same mapping db_writer.py uses, plus Subsource Dashboard, which is
# written by update_subsource.py instead of the main pipeline.
TAB_TABLE_MAP = dict(SHEET_TABLE_MAP)
TAB_TABLE_MAP["Subsource Dashboard"] = "subsource_dashboard"


def _df_to_sheet_values(df: pd.DataFrame):
    """[header_row, *data_rows] -- the same shape Google Sheets'
    getDataRange().getValues() returns, so Code.gs can't tell the
    difference between a real sheet tab and this API."""
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].astype(str)
    # route through pandas' own JSON encoder (handles NaN/Timestamps/
    # numpy types correctly) instead of hand-rolling type conversion
    split = json.loads(df.to_json(orient="split", date_format="iso"))
    return [split["columns"]] + split["data"]


@app.get("/dashboard/{month}")
def get_month_dashboard(month: str):
    """
    GET /dashboard/july
    Returns { data: { "Enhanced Branch Report": [[...]], ... } } for
    every tab in EXPECTED_TABS. Since your monthly Postgres tables only
    ever hold the CURRENT cycle's data (each upload replaces the last),
    this always returns whatever was most recently uploaded -- the
    {month} in the URL doesn't select a specific month's history, it's
    just there because that's the URL Code.gs already calls. If the
    month in the URL doesn't match the most recent upload's actual
    cycle, `warning` tells you so instead of failing silently.
    """
    data = {}
    for tab in EXPECTED_TABS:
        table_name = TAB_TABLE_MAP.get(tab)
        if not table_name:
            data[tab] = []
            continue
        try:
            df = pd.read_sql(f'SELECT * FROM "{table_name}"', engine)
            df = df.drop(columns=[c for c in ("_upload_id", "_run_timestamp") if c in df.columns])
            data[tab] = _df_to_sheet_values(df)
        except Exception as e:
            # Table doesn't exist yet (e.g. nothing uploaded this cycle) --
            # return an empty tab instead of failing the whole dashboard.
            data[tab] = []

    warning = None
    try:
        with engine.connect() as conn:
            latest = conn.execute(text("""
                SELECT cycle_month, cycle_year FROM upload_history
                ORDER BY run_timestamp DESC LIMIT 1
            """)).fetchone()
        if latest and latest.cycle_month and latest.cycle_month.lower() != month.lower():
            warning = f"Requested '{month}' but the most recent upload was {latest.cycle_month} {latest.cycle_year}."
    except Exception:
        pass

    return JSONResponse({"success": True, "month_requested": month, "warning": warning, "data": data})


# ============================================================
# /daily-tracker/{month} — same idea as /dashboard/{month}, but for
# the Overall_Dashboard + Track_* tables daily_tracker.py writes.
# Point Code.gs's DAILY_TRACKER_SHEETS at this instead of a Google
# Sheet ID once you want a given month live-synced from Neon.
# ============================================================

DAILY_TRACKER_TABS = [
    "Overall_Dashboard", "Track_PR", "Track_WhatsApp",
    "Track_AIMeeting", "Track_Website", "Track_LiveChat", "Track_RDMPL",
]


@app.get("/daily-tracker/{month}")
def get_daily_tracker(month: str):
    """
    GET /daily-tracker/july
    Same {data: {tab: [[...]]}} shape as /dashboard/{month}, sourced
    from daily_tracker.py's tables instead of the monthly ones.
    """
    data = {}
    for tab in DAILY_TRACKER_TABS:
        table_name = DAILY_TABLE_MAP.get(tab)
        if not table_name:
            data[tab] = []
            continue
        try:
            df = pd.read_sql(f'SELECT * FROM "{table_name}"', engine)
            data[tab] = _df_to_sheet_values(df)
        except Exception:
            data[tab] = []
    return JSONResponse({"success": True, "month_requested": month, "data": data})