import pandas as pd
from database import engine

# Dashboard name -> PostgreSQL table
TABLES = {
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
}


def get_dashboard_data():
    """
    Reads every dashboard table from Neon and returns
    the exact JSON structure expected by Apps Script.
    """

    dashboard = {}

    for sheet_name, table_name in TABLES.items():

        try:

            df = pd.read_sql(f'SELECT * FROM "{table_name}"', engine)

            # Remove internal columns
            for col in ["_upload_id", "_run_timestamp"]:
                if col in df.columns:
                    df.drop(columns=[col], inplace=True)

            # Convert NaN → None
            df = df.where(pd.notnull(df), None)

            dashboard[sheet_name] = df.to_dict(orient="records")

        except Exception as e:

            print(f"Error reading {table_name}: {e}")

            dashboard[sheet_name] = []

    return dashboard