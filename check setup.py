"""
check_setup.py
-----------------
Run this in PyCharm (or `python check_setup.py` from a terminal, in the
same folder as database.py) to sanity-check everything we've touched:
branch_targets, whether enhanced_branch_report looks fresh, when the
pipeline last ran, and whether daily_snapshots actually has more than
one date in it.

No arguments needed. Just run it and paste me the output.
"""

import pandas as pd
from database import engine


def section(title):
    print("\n" + "=" * 60)
    print(title)
    print("=" * 60)


def main():
    # ------------------------------------------------------------
    # 1. Does branch_targets exist, and does it look sane?
    # ------------------------------------------------------------
    section("1. branch_targets table")
    try:
        df = pd.read_sql('SELECT * FROM branch_targets ORDER BY "Branch"', engine)
        print(f"✅ Table exists, {len(df)} rows.\n")
        print(df.to_string(index=False))

        dupes = df["Branch"][df["Branch"].duplicated()].tolist()
        if dupes:
            print(f"\n⚠️  Duplicate branch names found: {dupes}")
    except Exception as e:
        print(f"❌ Could not read branch_targets: {e}")
        print("   -> Run: python update_branch_targets.py branch_targets.csv")

    # ------------------------------------------------------------
    # 2. When did the pipeline last actually finish successfully?
    # ------------------------------------------------------------
    section("2. Last successful pipeline run (upload_history)")
    try:
        hist = pd.read_sql(
            'SELECT * FROM upload_history ORDER BY 1 DESC LIMIT 5', engine
        )
        print(hist.to_string(index=False))
    except Exception as e:
        print(f"❌ Could not read upload_history: {e}")

    # ------------------------------------------------------------
    # 3. Does enhanced_branch_report look like it's using the new
    #    branch_targets-driven numbers, or old stale/duplicated ones?
    # ------------------------------------------------------------
    section("3. enhanced_branch_report freshness check")
    try:
        ebr = pd.read_sql('SELECT * FROM enhanced_branch_report', engine)
        print(f"{len(ebr)} rows.")

        if "branch_targets" in dir() or True:
            try:
                bt = pd.read_sql('SELECT "Branch", "Overall Target" FROM branch_targets', engine)
                bt_map = dict(zip(bt["Branch"], bt["Overall Target"]))
                mismatches = []
                for _, row in ebr.iterrows():
                    b = row.get("Branch")
                    if b in bt_map and row.get("Overall Target") not in (None, bt_map[b]):
                        mismatches.append((b, row.get("Overall Target"), bt_map[b]))
                if mismatches:
                    print(f"\n⚠️  {len(mismatches)} branch(es) in enhanced_branch_report do NOT "
                          f"match branch_targets -- this table is stale, re-run the pipeline:")
                    for b, got, expected in mismatches[:10]:
                        print(f"   {b}: report has {got}, branch_targets has {expected}")
                else:
                    print("✅ Overall Target values match branch_targets -- this table is fresh.")
            except Exception:
                pass

        odd_branches = [b for b in ebr.get("Branch", []) if b and b.lower() in
                        ("ahemdabad", "bnagalore")]
        if odd_branches:
            print(f"\n⚠️  Found known typo branch names still present: {odd_branches} "
                  f"-- this table predates the fix, re-run the pipeline.")
    except Exception as e:
        print(f"❌ Could not read enhanced_branch_report: {e}")

    # ------------------------------------------------------------
    # 4. daily_snapshots -- is history actually accumulating?
    # ------------------------------------------------------------
    section("4. daily_snapshots history")
    try:
        snaps = pd.read_sql(
            'SELECT "Snapshot Date", COUNT(*) AS rows '
            'FROM daily_snapshots GROUP BY 1 ORDER BY 1', engine
        )
        print(f"{len(snaps)} distinct date(s) in daily_snapshots:\n")
        print(snaps.to_string(index=False))
        if len(snaps) <= 1:
            print("\n⚠️  Only one date present. Either this is your first run, or you "
                  "haven't used the 'Backfill Daily Tracker history' checkbox yet.")
    except Exception as e:
        print(f"❌ Could not read daily_snapshots: {e}")

    section("Done")


if __name__ == "__main__":
    main()
