"""
db_writer.py
------------
Takes the exact `sheets` dict your notebook already builds in the
"FINAL EXCEL EXPORT" cell and writes each dataframe to a Neon Postgres
table instead of (or in addition to) the .xlsx file.

Nothing about your calculation logic changes. This only replaces the
destination of the final dataframes.
"""

import json
import uuid
from datetime import datetime

import pandas as pd
from sqlalchemy import text

from database import engine

# Maps the dict keys you already use in the notebook's `sheets = {...}`
# block to stable, lowercase, snake_case Postgres table names.
SHEET_TABLE_MAP = {
    "Enhanced Branch Report": "enhanced_branch_report",
    "Enhanced University Report": "enhanced_university_report",
    "Overall Dashboard": "overall_dashboard",
    "PR Dashboard": "pr_dashboard",
    "Whatsapp Dashboard": "whatsapp_dashboard",
    "AI Meeting Dashboard": "ai_meeting_dashboard",
    "Website Dashboard": "website_dashboard",
    "Live Chat Dashboard": "live_chat_dashboard",
    "Website Campaign Dashboard": "website_campaign_dashboard",
    "RDMPL Dashboard": "rdmpl_dashboard",
    "Product Dashboard": "product_dashboard",
    # "Subsource Dashboard" is in your .xlsx but is not produced anywhere
    # in the notebook, so it isn't wired up yet. Tell me where that sheet
    # comes from and I'll add it here too.
}


def _excel_safe(df: pd.DataFrame) -> pd.DataFrame:
    """Same text-cleaning your notebook already does before export,
    reused here so the DB copy matches the Excel copy exactly."""
    df = df.copy()
    for col in df.select_dtypes(include="object").columns:
        df[col] = (
            df[col]
            .fillna("")
            .astype(str)
            .str.replace(r"[\n\r\t]", " ", regex=True)
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
        )
    return df


def ensure_upload_history_table():
    with engine.begin() as conn:
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS upload_history (
                upload_id UUID PRIMARY KEY,
                run_timestamp TIMESTAMP NOT NULL,
                cycle_month TEXT,
                cycle_year INTEGER,
                raw_file TEXT,
                enrolled_file TEXT,
                targets_file TEXT,
                row_counts JSONB
            );
        """))


def save_report_to_postgres(
    sheets: dict,
    cycle_month: str = None,
    cycle_year: int = None,
    raw_file: str = None,
    enrolled_file: str = None,
    targets_file: str = None,
    if_exists: str = "replace",
):
    """
    Writes every dataframe in `sheets` to its matching Postgres table.

    if_exists="replace" (default) mirrors what your notebook does today:
    each run overwrites the table with the latest full snapshot, same as
    overwriting the .xlsx file. Switch to "append" later once you want to
    keep a history of every run instead of only the latest snapshot -- at
    that point add a run_date/upload_id column to each frame so rows from
    different runs stay distinguishable.
    """
    ensure_upload_history_table()

    upload_id = uuid.uuid4()
    run_timestamp = datetime.utcnow()
    row_counts = {}

    for sheet_name, df in sheets.items():
        if sheet_name not in SHEET_TABLE_MAP:
            print(f"⚠️  Skipping '{sheet_name}': no table mapping defined for it yet.")
            continue

        table_name = SHEET_TABLE_MAP[sheet_name]
        clean_df = _excel_safe(df)

        # tag every row with which run produced it, so future "append" mode
        # or debugging can trace a row back to its upload
        clean_df = clean_df.copy()
        clean_df["_upload_id"] = str(upload_id)
        clean_df["_run_timestamp"] = run_timestamp

        clean_df.to_sql(table_name, engine, if_exists=if_exists, index=False)
        row_counts[table_name] = len(clean_df)
        print(f"✅ {sheet_name} -> {table_name} ({len(clean_df)} rows)")

    with engine.begin() as conn:
        conn.execute(
            text("""
                INSERT INTO upload_history
                    (upload_id, run_timestamp, cycle_month, cycle_year,
                     raw_file, enrolled_file, targets_file, row_counts)
                VALUES
                    (:upload_id, :run_timestamp, :cycle_month, :cycle_year,
                     :raw_file, :enrolled_file, :targets_file, :row_counts)
            """),
            {
                "upload_id": str(upload_id),
                "run_timestamp": run_timestamp,
                "cycle_month": cycle_month,
                "cycle_year": cycle_year,
                "raw_file": raw_file,
                "enrolled_file": enrolled_file,
                "targets_file": targets_file,
                "row_counts": json.dumps(row_counts),
            },
        )

    print(f"\n📝 Logged run {upload_id} to upload_history")
    return upload_id
