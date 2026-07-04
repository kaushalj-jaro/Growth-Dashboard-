"""
database.py
------------
Single source of truth for the Neon Postgres connection.

SECURITY: never hardcode your connection string in this file or paste it
into a chat. Set it as an environment variable instead:

    Windows (PowerShell):
        setx DATABASE_URL "postgresql://user:password@ep-xxxx.neon.tech/dbname?sslmode=require"
        # close and reopen the terminal after setx

    Mac/Linux:
        export DATABASE_URL="postgresql://user:password@ep-xxxx.neon.tech/dbname?sslmode=require"

    Or put it in a local .env file (never commit this file to git):
        DATABASE_URL=postgresql://user:password@ep-xxxx.neon.tech/dbname?sslmode=require
"""

import os
from sqlalchemy import create_engine

# Loads a local .env file if python-dotenv is installed (optional convenience)
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

DATABASE_URL = os.environ.get("DATABASE_URL")

if not DATABASE_URL:
    raise RuntimeError(
        "DATABASE_URL is not set. Set it as an environment variable "
        "(see the comment at the top of database.py) before running the pipeline."
    )

# Neon requires SSL. psycopg2 needs the driver named explicitly for SQLAlchemy.
if DATABASE_URL.startswith("postgresql://") and "+psycopg2" not in DATABASE_URL:
    DATABASE_URL = DATABASE_URL.replace("postgresql://", "postgresql+psycopg2://", 1)

# pool_pre_ping avoids "server closed the connection" errors from Neon's
# autosuspend feature (Neon free tier pauses idle databases).
engine = create_engine(DATABASE_URL, pool_pre_ping=True)


def test_connection():
    """Quick sanity check you can run on its own: python database.py"""
    with engine.connect() as conn:
        from sqlalchemy import text
        result = conn.execute(text("SELECT version();"))
        print("✅ Connected to Neon Postgres")
        print(result.scalar())


if __name__ == "__main__":
    test_connection()
