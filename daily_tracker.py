"""
daily_tracker.py
-----------------
Your separate day-over-day tracker (Notebook 1 + Notebook 2 + the
Overall_Dashboard/Track_* workbook), ported to write into Postgres
instead of a growing Excel file. process_reports.py (your monthly
cycle pipeline) is untouched -- this is a second, independent module
that runs on the same daily Raw/Enrolled CSVs.

KEY CHANGE FROM YOUR ORIGINAL CODE: SNAPSHOT_DATE no longer comes from
parsing the file name. Instead:
  - the raw-side date comes from the max "Created On" in the Raw CSV
  - the enrolled-side date comes from the max "Modified On" in the
    Enrolled CSV (since DER Month is only a coarse month bucket, not
    a usable per-day anchor)
  - SNAPSHOT_DATE = the later of the two

ONE NAMING COLLISION, flagged rather than silently avoided: your old
workbook has a sheet called "Overall_Dashboard", and your MONTHLY
pipeline (process_reports.py / db_writer.py) already writes a table
called overall_dashboard. Writing this dashboard under the same name
would silently overwrite your monthly numbers with daily-tracker
numbers. I renamed just this one table to daily_overall_dashboard.
Every other table below uses your exact original sheet name.
"""

import calendar
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
from sqlalchemy import text

from database import engine

# ============================================================
# CONSTANTS (copied from your daily tracker notebook, unchanged)
# ============================================================

UNIVERSITY_STANDARDIZATION = {
    'Symbiosis School for Online and Digital Learning': 'Symbiosis School For Online And Digital Learning',
    'Manipal Unniversity, Jaipur': 'Manipal University, Jaipur',
    'MUJ': 'Manipal University, Jaipur',
    'MGR': 'Dr. M.G.R. University',
    'Amrita': 'Amrita Vishwa Vidyapeetham',
    'IIM Online': 'IIM',
}

BRANCH_STANDARDIZATION = {
    'Pune1': 'Pune 1',
    'Thane 1': 'Thane',
    'Noida': 'Noida 1',
    'Pune': 'Pune 1',
    'Thane 2': 'Thane',
    'Goregoan': 'Goregaon',
}

SOURCE_STANDARDIZATION = {
    'whatsapp': 'WhatsApp', 'whatsapp reference': 'WhatsApp', 'whatsapp ref': 'WhatsApp',
    'whatsapp lead': 'WhatsApp', 'whatsapp leads': 'WhatsApp', 'whatsappreference': 'WhatsApp',
    'whats app': 'WhatsApp', 'wa': 'WhatsApp', 'wa ref': 'WhatsApp', 'wa reference': 'WhatsApp',
    'pr': 'PR', 'pr story': 'PR', 'prstory': 'PR', 'prstory campaign': 'PR',
    'prchat': 'PR', 'prchat campaign': 'PR',
    'ai meeting': 'AI Meeting', 'career counselling': 'AI Meeting', 'career counseling': 'AI Meeting', 'superbot': 'AI Meeting',
    'website': 'Website', 'inbound call': 'Website', 'inbound calls': 'Website', 'inbound phone call': 'Website',
    'inbound phone calls': 'Website', 'inbound enquiry': 'Website', 'inbound inquiry': 'Website', 'web': 'Website',
    'organic website': 'Website', 'website lead': 'Website',
    'live chat': 'Live Chat', 'livechat': 'Live Chat', 'live_chat': 'Live Chat', 'chat': 'Live Chat', 'live-chat': 'Live Chat', 'lc': 'Live Chat',
    'rdmpl': 'RDMPL', 'rdmpl lead': 'RDMPL', 'rdmpl campaign': 'RDMPL',
}

FINAL_NORM = {
    'live chat': 'Live Chat', 'livechat': 'Live Chat', 'live-chat': 'Live Chat', 'live_chat': 'Live Chat',
    'lc': 'Live Chat', 'chat': 'Live Chat', 'superbot': 'AI Meeting', 'whatsapp': 'WhatsApp',
    'pr': 'PR', 'website': 'Website', 'ai meeting': 'AI Meeting', 'rdmpl': 'RDMPL',
}

JUNK = ['Invalid', 'Disqualified', 'DND', 'Lead Not Enquired']
WORKABLE = ['Callback', 'Eligible for DLP', 'Future Follow Up', 'Marketing Qualified', 'Not Contactable', 'Reborn', 'Attempted']
PROSPECT = ['Prospect', 'Meeting Done', 'Meeting Scheduled', 'Registered', 'Registered Lead', 'Opportunity']
FRESH = ['Fresh']

SOURCES = ['PR', 'WhatsApp', 'AI Meeting', 'Website', 'Live Chat', 'RDMPL', 'Others']
FUNNEL_COLS = ['Delivered', 'Workable', 'Prospect', 'Fresh', 'Junk', 'Current Adm', 'Spillover', 'Total Adm']

PV_UNIVERSITIES = [str(x).strip() for x in [
    'IIM Ahmedabad', 'IIM Bangalore', 'IIM Calcutta', 'IIM Indore', 'IIM Kozhikode',
    'IIM Lucknow', 'IIM Mumbai', 'IIM Nagpur', 'IIM Raipur', 'IIM Ranchi', 'IIM Rohtak',
    'IIM Sambalpur', 'IIM Shillong', 'IIM Sirmaur', 'IIM Tiruchirappalli', 'IIM Trichy',
    'IIM Udaipur', 'IIM Visakhapatnam', 'IIT Bombay', 'IIT Delhi', 'IIT Guwahati',
    'IIT Kanpur', 'IIT Kharagpur', 'IIT Madras', 'IIT Roorkee', 'XLRI, Jamshedpur'
]]

# Daily benchmark targets per channel (from your Track_* sheets)
CHANNEL_BENCHMARKS = {
    'WhatsApp':   {'p_tgt': 20, 'a_tgt': 5,  'pfx': 'WhatsApp_'},
    'Live Chat':  {'p_tgt': 30, 'a_tgt': 8,  'pfx': 'Live_Chat_'},
    'AI Meeting': {'p_tgt': 30, 'a_tgt': 6,  'pfx': 'AI_Meeting_'},
    'PR':         {'p_tgt': 50, 'a_tgt': 15, 'pfx': 'PR_'},
    'Website':    {'p_tgt': 40, 'a_tgt': 12, 'pfx': 'Website_'},
    'RDMPL':      {'p_tgt': 25, 'a_tgt': 5,  'pfx': 'RDMPL_'},
}

# sheet name -> Postgres table name. Everything uses your exact original
# sheet name except Overall_Dashboard (see module docstring for why).
TABLE_MAP = {
    'Daily_Snapshots': 'daily_snapshots',
    'Daily_Summary': 'Daily_Summary',
    'Today_vs_Yesterday': 'Today_vs_Yesterday',
    'Branch_Today': 'Branch_Today',
    'Source_Today': 'Source_Today',
    'Branch_Delivered_Trend': 'Branch_Delivered_Trend',
    'Branch_Admissions_Trend': 'Branch_Admissions_Trend',
    'University_Trend': 'University_Trend',
    'Source_Trend': 'Source_Trend',
    'Source_Admissions_Trend': 'Source_Admissions_Trend',
    'Full_Delta_Detail': 'Full_Delta_Detail',
    'Month_To_Date_Summary': 'Month_To_Date_Summary',
    # 'Full_History' is intentionally skipped -- it's identical to Daily_Snapshots
    'Overall_Dashboard': 'daily_overall_dashboard',  # renamed to avoid clobbering the monthly table
    'Track_WhatsApp': 'Track_WhatsApp',
    'Track_LiveChat': 'Track_LiveChat',
    'Track_AIMeeting': 'Track_AIMeeting',
    'Track_PR': 'Track_PR',
    'Track_Website': 'Track_Website',
    'Track_RDMPL': 'Track_RDMPL',
}


def _write(df: pd.DataFrame, table_name: str):
    df.to_sql(table_name, engine, if_exists='replace', index=False)
    print(f"✅ -> {table_name} ({len(df)} rows)")


def clean_dataset(df, is_enrolled=False):
    df.columns = df.columns.str.strip()
    if 'Created On' in df.columns:
        df['Created On'] = pd.to_datetime(df['Created On'], dayfirst=True, errors='coerce')

    text_cols = ['Branch', 'Admission Branch', 'University', 'OFS 1 - Primary University', 'Opportunity Source', 'Stage', 'DER Month', 'Product', 'Opportunity Id']
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].fillna('').astype(str).str.strip().str.replace(r'[\n\r\t]', ' ', regex=True).str.replace(r'\s+', ' ', regex=True)

    if 'Branch' in df.columns:
        df['Branch'] = df['Branch'].replace(BRANCH_STANDARDIZATION)
    if 'Admission Branch' in df.columns:
        df['Admission Branch'] = df['Admission Branch'].replace(BRANCH_STANDARDIZATION)
    if 'University' in df.columns:
        df['University'] = df['University'].replace(UNIVERSITY_STANDARDIZATION)
    if 'OFS 1 - Primary University' in df.columns:
        df['OFS 1 - Primary University'] = df['OFS 1 - Primary University'].replace(UNIVERSITY_STANDARDIZATION)

    if 'Opportunity Source' in df.columns:
        df['Opportunity Source'] = df['Opportunity Source'].str.lower().replace(SOURCE_STANDARDIZATION)
        df['Opportunity Source'] = df['Opportunity Source'].replace(FINAL_NORM)

    if 'Branch' in df.columns:
        df = df[(df['Branch'] != '') & (df['Branch'].str.lower() != 'nan') & (df['Branch'].isna() == False)].copy()
    if is_enrolled and 'Admission Branch' in df.columns:
        df = df[(df['Admission Branch'] != '') & (df['Admission Branch'].str.lower() != 'nan') & (df['Admission Branch'].isna() == False)].copy()

    if 'Opportunity Id' in df.columns:
        df = df.drop_duplicates(subset=['Opportunity Id'])

    return df


MONTH_TO_NUM = {
    'january': 1, 'february': 2, 'march': 3, 'april': 4, 'may': 5, 'june': 6,
    'july': 7, 'august': 8, 'september': 9, 'october': 10, 'november': 11, 'december': 12,
    'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'jun': 6, 'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12,
}


def to_int_safe(val):
    if pd.isna(val) or str(val).strip() == '':
        return -1
    cleaned = str(val).split('.')[0].strip().lower()
    if cleaned in MONTH_TO_NUM:
        return MONTH_TO_NUM[cleaned]
    try:
        return int(float(cleaned))
    except Exception:
        return -1


def _prepare_data(raw_path: str, enrolled_path: str):
    """
    Load, clean, segment, and build combos ONCE. Both a normal live run
    and a backfill run share this exact same preparation -- there's only
    one code path for "what does the data look like", so the two modes
    can never silently drift apart from each other.
    """
    df_raw = pd.read_csv(raw_path, low_memory=False)
    df_enrolled = pd.read_csv(enrolled_path, low_memory=False)
    df_raw.columns = df_raw.columns.str.strip()
    df_enrolled.columns = df_enrolled.columns.str.strip()

    yesterday = datetime.now().date() - timedelta(days=1)

    if 'Created On' in df_raw.columns:
        _raw_dates = pd.to_datetime(df_raw['Created On'], dayfirst=True, errors='coerce')
        RAW_SNAPSHOT_DATE = _raw_dates.max().date() if _raw_dates.notna().any() else yesterday
    else:
        print("⚠️ 'Created On' not found in raw file -- falling back to yesterday's date.")
        RAW_SNAPSHOT_DATE = yesterday

    if 'Modified On' in df_enrolled.columns:
        _enr_dates = pd.to_datetime(df_enrolled['Modified On'], dayfirst=True, errors='coerce')
        ENROLLED_SNAPSHOT_DATE = _enr_dates.max().date() if _enr_dates.notna().any() else RAW_SNAPSHOT_DATE
    else:
        print("⚠️ 'Modified On' not found in enrolled file -- using the raw-side date instead.")
        ENROLLED_SNAPSHOT_DATE = RAW_SNAPSHOT_DATE

    SNAPSHOT_DATE = max(RAW_SNAPSHOT_DATE, ENROLLED_SNAPSHOT_DATE)

    CYCLE_MONTH = calendar.month_name[SNAPSHOT_DATE.month]
    CYCLE_YEAR = SNAPSHOT_DATE.year
    month_idx = SNAPSHOT_DATE.month
    CYCLE_START = f'{CYCLE_YEAR}-{month_idx:02d}-01 00:00:00'

    print("====================================================")
    print("📅 DAILY TRACKER — SNAPSHOT DATE (DATA-DRIVEN)")
    print("====================================================")
    print(f"Raw side (max 'Created On')      : {RAW_SNAPSHOT_DATE}")
    print(f"Enrolled side (max 'Modified On') : {ENROLLED_SNAPSHOT_DATE}")
    print(f"SNAPSHOT_DATE in use              : {SNAPSHOT_DATE}")

    df_raw = clean_dataset(df_raw, is_enrolled=False)
    df_enrolled = clean_dataset(df_enrolled, is_enrolled=True)

    df_raw['Segment'] = 'DLP'
    pv_mask = df_raw['University'].isin(PV_UNIVERSITIES) | df_raw['OFS 1 - Primary University'].isin(PV_UNIVERSITIES)
    df_raw.loc[pv_mask, 'Segment'] = 'PV'
    df_raw.loc[df_raw['University'] == 'Jaro Education', 'Segment'] = 'Free Courses'

    free_course_ids = set(df_raw[df_raw['Segment'] == 'Free Courses']['Opportunity Id'])
    df_enrolled['Segment'] = 'DLP'
    pv_mask_enrolled = df_enrolled['University'].isin(PV_UNIVERSITIES) | df_enrolled['OFS 1 - Primary University'].isin(PV_UNIVERSITIES)
    df_enrolled.loc[pv_mask_enrolled, 'Segment'] = 'PV'
    df_enrolled.loc[df_enrolled['Opportunity Id'].isin(free_course_ids), 'Segment'] = 'Free Courses'

    if 'Admission Branch' in df_enrolled.columns:
        df_enrolled['Final Branch'] = df_enrolled['Admission Branch']
    else:
        df_enrolled['Final Branch'] = df_enrolled['Branch']
    blank_br = (df_enrolled['Final Branch'] == '') | (df_enrolled['Final Branch'].isna())
    df_enrolled.loc[blank_br, 'Final Branch'] = df_enrolled.loc[blank_br, 'Branch']

    if 'OFS 1 - Primary University' in df_enrolled.columns:
        mask_enr_routed = (
            (df_enrolled['University'] == 'Any') &
            (df_enrolled['OFS 1 - Primary University'] != '') &
            (df_enrolled['OFS 1 - Primary University'] != 'Any') &
            (df_enrolled['OFS 1 - Primary University'].str.lower() != 'nan')
        )
        df_enrolled.loc[mask_enr_routed, 'University'] = df_enrolled.loc[mask_enr_routed, 'OFS 1 - Primary University']

    df_raw.loc[df_raw['Segment'] == 'Free Courses', 'University'] = 'Jaro Education'
    df_raw.loc[df_raw['Segment'] == 'Free Courses', 'OFS 1 - Primary University'] = ''
    df_enrolled.loc[df_enrolled['Opportunity Id'].isin(free_course_ids), 'University'] = 'Jaro Education'

    # Month-scope the raw leads: only leads CREATED this month count,
    # same principle admissions already used. Rows with an unparseable
    # date are kept rather than dropped, but logged.
    if 'Created On' in df_raw.columns:
        _before = len(df_raw)
        _undated = df_raw['Created On'].isna().sum()
        df_raw = df_raw[df_raw['Created On'].isna() | (df_raw['Created On'] >= pd.to_datetime(CYCLE_START))].copy()
        _dropped = _before - len(df_raw)
        print(f"📅 Month-scoped raw leads to {CYCLE_MONTH} {CYCLE_YEAR}: kept {len(df_raw)}/{_before} rows "
              f"({_dropped} from prior months excluded, {_undated} with no parseable date kept as-is)")

    isolated_lc_raw = df_raw[df_raw['Opportunity Source'].str.lower() == 'live chat'].copy()
    isolated_lc_enrolled = df_enrolled[df_enrolled['Opportunity Source'].str.lower() == 'live chat'].copy()

    for frame in [df_raw, df_enrolled, isolated_lc_raw, isolated_lc_enrolled]:
        if frame is not None and not frame.empty:
            if 'DER Month' in frame.columns:
                frame['DER_Month_Int'] = frame['DER Month'].apply(to_int_safe)
            if 'DER Year' in frame.columns:
                frame['DER_Year_Int'] = frame['DER Year'].apply(to_int_safe)

    target_month_int = MONTH_TO_NUM.get(str(CYCLE_MONTH).strip().lower(), SNAPSHOT_DATE.month)

    df_raw['Combo_University'] = df_raw['University']
    mask_routed = (df_raw['University'] == 'Any') & (df_raw['OFS 1 - Primary University'] != '')
    df_raw.loc[mask_routed, 'Combo_University'] = df_raw.loc[mask_routed, 'OFS 1 - Primary University']
    mask_blank_any = (df_raw['University'] == 'Any') & (df_raw['OFS 1 - Primary University'] == '')
    df_raw.loc[mask_blank_any, 'Combo_University'] = 'Any'

    delivered_combos = df_raw[['Branch', 'Combo_University', 'Segment']].copy().rename(columns={'Combo_University': 'University'})
    admission_combos = df_enrolled[['Final Branch', 'University', 'Segment']].copy().rename(columns={'Final Branch': 'Branch'})
    combos = pd.concat([delivered_combos, admission_combos], ignore_index=True)
    combos = combos[(combos['Branch'] != '') & (combos['University'] != '') & (combos['Segment'] != '')].drop_duplicates(subset=['Branch', 'University', 'Segment'])

    valid_combos = []
    for _, row in combos.iterrows():
        branch, university, segment = row['Branch'], row['University'], row['Segment']
        if university == 'Any':
            raw_mask = (df_raw['University'] == 'Any') & (df_raw['OFS 1 - Primary University'] == '')
        else:
            raw_mask = (df_raw['University'] == university) | ((df_raw['University'] == 'Any') & (df_raw['OFS 1 - Primary University'] == university))

        has_base_data = len(df_raw[(df_raw['Branch'] == branch) & (df_raw['Segment'] == segment) & raw_mask]) > 0 or \
                         len(df_enrolled[(df_enrolled['Final Branch'] == branch) & (df_enrolled['Segment'] == segment) & (df_enrolled['University'] == university)]) > 0
        has_rdmpl_data = (university == 'Any') and (len(df_raw[(df_raw['Branch'] == branch) & (df_raw['Opportunity Source'].str.upper() == 'RDMPL')]) > 0)

        if has_base_data or has_rdmpl_data:
            valid_combos.append({'Branch': branch, 'University': university, 'Segment': segment})

    combos = pd.DataFrame(valid_combos).sort_values(['Segment', 'Branch', 'University']).reset_index(drop=True)

    return {
        'df_raw': df_raw, 'df_enrolled': df_enrolled,
        'isolated_lc_raw': isolated_lc_raw, 'isolated_lc_enrolled': isolated_lc_enrolled,
        'combos': combos, 'target_month_int': target_month_int,
        'CYCLE_START': CYCLE_START, 'CYCLE_MONTH': CYCLE_MONTH, 'CYCLE_YEAR': CYCLE_YEAR,
        'SNAPSHOT_DATE': SNAPSHOT_DATE, 'RAW_SNAPSHOT_DATE': RAW_SNAPSHOT_DATE,
        'ENROLLED_SNAPSHOT_DATE': ENROLLED_SNAPSHOT_DATE,
    }


def _get_metrics(prep, branch, university, segment, source_filter=None, as_of_date=None):
    """
    Same metric engine as before, plus an optional as_of_date cutoff used
    only by backfill: when set, both raw leads and admissions are capped
    to "Created On <= end of as_of_date", reconstructing what the numbers
    would have looked like on that historical day. When as_of_date is
    None (the normal live run), behavior is identical to before -- there's
    nothing beyond today in the file anyway.
    """
    if source_filter == 'Live Chat':
        working_raw, working_enr = prep['isolated_lc_raw'], prep['isolated_lc_enrolled']
    else:
        working_raw, working_enr = prep['df_raw'], prep['df_enrolled']

    if university == 'Any':
        if source_filter == 'AI Meeting':
            raw_mask = (working_raw['Opportunity Source'] == 'AI Meeting') & \
                       (working_raw['University'] == 'Any') & \
                       (working_raw['OFS 1 - Primary University'].isin(['', 'Any']) | working_raw['OFS 1 - Primary University'].isna())
        else:
            raw_mask = (working_raw['University'] == 'Any') & ((working_raw['OFS 1 - Primary University'] == '') | (working_raw['OFS 1 - Primary University'] == 'Any') | (working_raw['OFS 1 - Primary University'].isna()))
    else:
        raw_mask = (working_raw['University'] == university) | ((working_raw['University'] == 'Any') & (working_raw['OFS 1 - Primary University'] == university))

    if source_filter == 'RDMPL':
        raw_temp = working_raw[(working_raw['Branch'] == branch) & (working_raw['University'] == university) & (working_raw['Opportunity Source'] == 'RDMPL')].copy()
    elif source_filter == 'Live Chat':
        raw_temp = working_raw[(working_raw['Branch'] == branch) & (working_raw['University'] == university) & (working_raw['Opportunity Source'] == 'Live Chat')].copy()
    else:
        raw_temp = working_raw[(working_raw['Branch'] == branch) & (working_raw['Segment'] == segment) & raw_mask].copy()
        if source_filter is None:
            raw_temp = raw_temp[raw_temp['Opportunity Source'] != 'Live Chat']
        else:
            raw_temp = raw_temp[raw_temp['Opportunity Source'] == source_filter]

    if 'Opportunity Id' in raw_temp.columns:
        raw_temp = raw_temp.drop_duplicates(subset=['Opportunity Id'])

    enr_temp = working_enr[(working_enr['Final Branch'] == branch) & (working_enr['Segment'] == segment) & (working_enr['DER_Month_Int'] == prep['target_month_int']) & (working_enr['University'] == university)].copy()
    if source_filter is None:
        enr_temp = enr_temp[enr_temp['Opportunity Source'] != 'Live Chat']
    else:
        enr_temp = enr_temp[enr_temp['Opportunity Source'] == source_filter]
    if 'Opportunity Id' in enr_temp.columns:
        enr_temp = enr_temp.drop_duplicates(subset=['Opportunity Id'])

    if as_of_date is not None:
        cutoff = pd.to_datetime(as_of_date) + pd.Timedelta(days=1)  # end of that day
        if 'Created On' in raw_temp.columns:
            raw_temp = raw_temp[raw_temp['Created On'].isna() | (raw_temp['Created On'] < cutoff)]
        enr_temp = enr_temp[pd.to_datetime(enr_temp['Created On'], dayfirst=True, errors='coerce') < cutoff]

    delivered = len(raw_temp)
    workable = len(raw_temp[raw_temp['Stage'].isin(WORKABLE)])
    prospect = len(raw_temp[raw_temp['Stage'].isin(PROSPECT)])
    fresh = len(raw_temp[raw_temp['Stage'].isin(FRESH)])
    junk = len(raw_temp[raw_temp['Stage'].isin(JUNK)])

    enr_temp = enr_temp.copy()
    enr_temp['Created On'] = pd.to_datetime(enr_temp['Created On'], dayfirst=True, errors='coerce')
    valid_dates = enr_temp.dropna(subset=['Created On'])
    failed_dates_count = len(enr_temp) - len(valid_dates)

    current_adm = len(valid_dates[valid_dates['Created On'] >= pd.to_datetime(prep['CYCLE_START'])]) + failed_dates_count
    spillover = len(valid_dates[valid_dates['Created On'] < pd.to_datetime(prep['CYCLE_START'])])
    total_adm = current_adm + spillover

    cvr = round((total_adm / delivered) * 100, 2) if delivered > 0 else 0
    junk_pct = round((junk / delivered) * 100, 2) if delivered > 0 else 0

    return {
        'Delivered': delivered, 'Workable': workable, 'Prospect': prospect, 'Fresh': fresh, 'Junk': junk,
        'Current Adm': current_adm, 'Spillover': spillover, 'Total Adm': total_adm, 'CVR %': cvr, 'Junk %': junk_pct
    }


def _build_snapshot_df(prep, as_of_date=None):
    """
    Builds one snapshot row per Branch x University x Segment, either for
    today (as_of_date=None) or for a historical cutoff (backfill).
    Tags every row with `_backfilled`: True only for a STRICTLY PAST cutoff
    (Workable/Prospect/Fresh/Junk there reflect today's current stage, not
    the actual historical stage). The day equal to SNAPSHOT_DATE itself is
    never tagged backfilled, since that IS the real live capture.
    """
    snapshot_date = as_of_date if as_of_date is not None else prep['SNAPSHOT_DATE']
    is_backfilled = as_of_date is not None and as_of_date < prep['SNAPSHOT_DATE']

    rows = []
    for _, combo in prep['combos'].iterrows():
        branch, university, segment = combo['Branch'], combo['University'], combo['Segment']
        channel_metrics = {src: _get_metrics(prep, branch, university, segment, source_filter=src, as_of_date=as_of_date) for src in SOURCES}
        overall = _get_metrics(prep, branch, university, segment, source_filter=None, as_of_date=as_of_date)

        row = {
            'Snapshot Date': str(snapshot_date), 'Branch': branch, 'University': university, 'Segment': segment,
            '_backfilled': is_backfilled,
            'Delivered': overall['Delivered'], 'Workable': overall['Workable'], 'Prospect': overall['Prospect'],
            'Fresh': overall['Fresh'], 'Junk': overall['Junk'], 'Current Adm': overall['Current Adm'],
            'Spillover': overall['Spillover'], 'Total Adm': overall['Total Adm'], 'CVR %': overall['CVR %'], 'Junk %': overall['Junk %'],
        }
        for src in SOURCES:
            m = channel_metrics[src]
            prefix = src.replace(' ', '_') + '_'
            row[prefix + 'Delivered'] = m['Delivered']
            row[prefix + 'Workable'] = m['Workable']
            row[prefix + 'Prospect'] = m['Prospect']
            row[prefix + 'Fresh'] = m['Fresh']
            row[prefix + 'Junk'] = m['Junk']
            row[prefix + 'Current_Adm'] = m['Current Adm']
            row[prefix + 'Spillover'] = m['Spillover']
            row[prefix + 'Total_Adm'] = m['Total Adm']
            row[prefix + 'CVR_pct'] = m['CVR %']
            row[prefix + 'Junk_pct'] = m['Junk %']
        rows.append(row)

    snap_df = pd.DataFrame(rows)

    dynamic_ai_target = len(prep['df_raw'][
        (prep['df_raw']['Opportunity Source'] == 'AI Meeting') &
        (as_of_date is None or (prep['df_raw']['Created On'].isna() | (prep['df_raw']['Created On'] < pd.to_datetime(as_of_date) + pd.Timedelta(days=1))))
    ])
    current_ai_total = snap_df['AI_Meeting_Delivered'].sum()
    if current_ai_total != dynamic_ai_target:
        discrepancy = dynamic_ai_target - current_ai_total
        any_idx = snap_df[snap_df['University'] == 'Any'].index
        if not any_idx.empty:
            snap_df.loc[any_idx[0], 'AI_Meeting_Delivered'] += discrepancy
            snap_df.loc[any_idx[0], 'AI_Meeting_Total_Adm'] = snap_df.loc[any_idx[0], 'AI_Meeting_Current_Adm'] + snap_df.loc[any_idx[0], 'AI_Meeting_Spillover']
            snap_df.loc[any_idx[0], 'Delivered'] += discrepancy

    return snap_df


def _upsert_snapshot(snap_df, snapshot_date):
    """Deletes any existing rows for this date, then inserts the fresh
    ones -- same-day reruns overwrite that day only, every other day in
    the history stays untouched. Also self-heals the table schema if a
    column (like the new `_backfilled` flag) doesn't exist yet on a
    table created by an older version of this script."""
    with engine.begin() as conn:
        table_exists = conn.execute(text("SELECT to_regclass('public.daily_snapshots')")).scalar()
        if table_exists:
            existing_cols = {row[0] for row in conn.execute(text("""
                SELECT column_name FROM information_schema.columns WHERE table_name = 'daily_snapshots'
            """))}
            for col in snap_df.columns:
                if col not in existing_cols:
                    dtype = snap_df[col].dtype
                    pg_type = ('BOOLEAN' if dtype == bool else
                               'DOUBLE PRECISION' if dtype.kind == 'f' else
                               'BIGINT' if dtype.kind in 'iu' else 'TEXT')
                    conn.execute(text(f'ALTER TABLE daily_snapshots ADD COLUMN IF NOT EXISTS "{col}" {pg_type}'))
            conn.execute(text('DELETE FROM daily_snapshots WHERE "Snapshot Date" = :d'), {"d": str(snapshot_date)})
    snap_df.to_sql('daily_snapshots', engine, if_exists='append', index=False)
    print(f"✅ -> daily_snapshots (+{len(snap_df)} rows for {snapshot_date}, older days untouched)")

def _rebuild_derived_tables():
    """Recomputes every derived table (Daily_Summary, trends, delta,
    MTD summary, Overall_Dashboard, Track_*) from the full daily_snapshots
    history currently in Postgres. Both a live run and a backfill call
    this once, after all their snapshot rows are written."""
    master = pd.read_sql('SELECT * FROM daily_snapshots', engine)
    master['Snapshot Date'] = pd.to_datetime(master['Snapshot Date'])
    dates_logged = sorted(master['Snapshot Date'].unique())
    latest_date = dates_logged[-1]
    prev_date = dates_logged[-2] if len(dates_logged) >= 2 else None
    has_prev = prev_date is not None
    today_label = str(pd.Timestamp(latest_date).date())
    prev_label = str(pd.Timestamp(prev_date).date()) if has_prev else 'N/A'

    today_data = master[master['Snapshot Date'] == latest_date].copy()
    yest_data = master[master['Snapshot Date'] == prev_date].copy() if has_prev else pd.DataFrame(columns=today_data.columns)

    # ---- Daily_Summary ----
    summary = master.groupby('Snapshot Date', as_index=False).agg(
        Total_Combos=('Branch', 'count'), Total_Delivered=('Delivered', 'sum'), Total_Workable=('Workable', 'sum'),
        Total_Prospect=('Prospect', 'sum'), Total_Fresh=('Fresh', 'sum'), Total_Junk=('Junk', 'sum'), Total_Adm=('Total Adm', 'sum'),
    )
    summary['Avg_CVR%'] = (summary['Total_Adm'] / summary['Total_Delivered'].replace(0, 1) * 100).round(2)
    _write(summary, TABLE_MAP['Daily_Summary'])

    # ============================================================
    # TREND TABLES
    # ============================================================
    master['Date_Str'] = master['Snapshot Date'].dt.strftime('%d-%b')

    branch_trend = master.groupby(['Snapshot Date', 'Branch'], as_index=False).agg(Delivered=('Delivered', 'sum'), Total_Adm=('Total Adm', 'sum'))
    branch_trend['Date_Str'] = branch_trend['Snapshot Date'].dt.strftime('%d-%b')
    branch_pivot_delivered = branch_trend.pivot_table(index='Branch', columns='Date_Str', values='Delivered', fill_value=0).reset_index()
    branch_pivot_adm = branch_trend.pivot_table(index='Branch', columns='Date_Str', values='Total_Adm', fill_value=0).reset_index()
    _write(branch_pivot_delivered, TABLE_MAP['Branch_Delivered_Trend'])
    _write(branch_pivot_adm, TABLE_MAP['Branch_Admissions_Trend'])

    uni_trend = master.groupby(['Snapshot Date', 'University'], as_index=False).agg(Delivered=('Delivered', 'sum'), Total_Adm=('Total Adm', 'sum'))
    uni_trend['Date_Str'] = uni_trend['Snapshot Date'].dt.strftime('%d-%b')
    uni_pivot_delivered = uni_trend.pivot_table(index='University', columns='Date_Str', values='Delivered', fill_value=0).reset_index()
    _write(uni_pivot_delivered, TABLE_MAP['University_Trend'])

    src_delivered_rows, src_admission_rows = [], []
    for src in SOURCES:
        src_del_col = f"{src.replace(' ', '_')}_Delivered"
        src_adm_col = f"{src.replace(' ', '_')}_Total_Adm"
        if src_del_col in master.columns:
            grp_del = master.groupby('Snapshot Date')[src_del_col].sum().reset_index()
            grp_del['Source'] = src
            grp_del = grp_del.rename(columns={src_del_col: 'Delivered'})
            src_delivered_rows.append(grp_del)
        if src_adm_col in master.columns:
            grp_adm = master.groupby('Snapshot Date')[src_adm_col].sum().reset_index()
            grp_adm['Source'] = src
            grp_adm = grp_adm.rename(columns={src_adm_col: 'Total_Adm'})
            src_admission_rows.append(grp_adm)

    if src_delivered_rows:
        source_trend_del_df = pd.concat(src_delivered_rows, ignore_index=True)
        source_trend_del_df['Date_Str'] = source_trend_del_df['Snapshot Date'].dt.strftime('%d-%b')
        source_pivot = source_trend_del_df.pivot_table(index='Source', columns='Date_Str', values='Delivered', fill_value=0).reset_index()
    else:
        source_pivot = pd.DataFrame(columns=['Source'])
    _write(source_pivot, TABLE_MAP['Source_Trend'])

    if src_admission_rows:
        source_trend_adm_df = pd.concat(src_admission_rows, ignore_index=True)
        source_trend_adm_df['Date_Str'] = source_trend_adm_df['Snapshot Date'].dt.strftime('%d-%b')
        source_pivot_adm = source_trend_adm_df.pivot_table(index='Source', columns='Date_Str', values='Total_Adm', fill_value=0).reset_index()
    else:
        source_pivot_adm = pd.DataFrame(columns=['Source'])
    _write(source_pivot_adm, TABLE_MAP['Source_Admissions_Trend'])

    # ============================================================
    # TODAY VS YESTERDAY — BRANCH LEVEL (-> Today_vs_Yesterday)
    # ============================================================
    non_metric_keys = ['Snapshot Date', 'Branch', 'University', 'Segment', 'Date_Str']
    all_horizontal_metrics = [c for c in today_data.columns if c not in non_metric_keys and not c.endswith('%') and 'pct' not in c]
    aggregation_rules = {m: 'sum' for m in all_horizontal_metrics}

    today_branch = today_data.groupby('Branch', as_index=False).agg(aggregation_rules).rename(columns={c: f"{c}_T" for c in all_horizontal_metrics})
    if has_prev:
        yest_branch = yest_data.groupby('Branch', as_index=False).agg(aggregation_rules).rename(columns={c: f"{c}_Y" for c in all_horizontal_metrics})
        compare = today_branch.merge(yest_branch, on='Branch', how='outer').fillna(0)
    else:
        compare = today_branch.copy()
        for c in all_horizontal_metrics:
            compare[f"{c}_Y"] = 0

    for m in all_horizontal_metrics:
        compare[f"{m}_Δ"] = compare[f"{m}_T"] - compare[f"{m}_Y"]

    compare['CVR%_Today'] = (compare['Total Adm_T'] / compare['Delivered_T'].replace(0, np.nan) * 100).round(2).fillna(0)
    compare['Junk%_Today'] = (compare['Junk_T'] / compare['Delivered_T'].replace(0, np.nan) * 100).round(2).fillna(0)

    runtime_sources = sorted(set(c.split('_Delivered')[0] for c in all_horizontal_metrics if c.endswith('_Delivered')))
    for src in runtime_sources:
        del_col, junk_col = f"{src}_Delivered", f"{src}_Junk"
        adm_col = f"{src}_Total_Adm"
        if f"{del_col}_T" in compare.columns:
            if f"{adm_col}_T" in compare.columns:
                compare[f"{src}_CVR%_Today"] = (compare[f"{adm_col}_T"] / compare[f"{del_col}_T"].replace(0, np.nan) * 100).round(2).fillna(0)
            if f"{junk_col}_T" in compare.columns:
                compare[f"{src}_Junk%_Today"] = (compare[f"{junk_col}_T"] / compare[f"{del_col}_T"].replace(0, np.nan) * 100).round(2).fillna(0)

    display_cols_order = ['Branch']
    for stage in FUNNEL_COLS:
        display_cols_order.extend([f"{stage}_Y", f"{stage}_T", f"{stage}_Δ"])
    display_cols_order.extend(['CVR%_Today', 'Junk%_Today'])
    for src in runtime_sources:
        for stage in FUNNEL_COLS:
            col = f"{src}_{stage.replace(' ', '_')}"
            if f"{col}_T" in compare.columns:
                display_cols_order.extend([f"{col}_Y", f"{col}_T", f"{col}_Δ"])
        if f"{src}_CVR%_Today" in compare.columns:
            display_cols_order.append(f"{src}_CVR%_Today")
        if f"{src}_Junk%_Today" in compare.columns:
            display_cols_order.append(f"{src}_Junk%_Today")

    display_cols_order = [c for c in display_cols_order if c in compare.columns]
    compare_display = compare[display_cols_order].copy()
    rename_map = {}
    for m in all_horizontal_metrics:
        rename_map[f"{m}_Y"] = f"{m} ({prev_label})"
        rename_map[f"{m}_T"] = f"{m} ({today_label})"
        rename_map[f"{m}_Δ"] = f"{m} Δ"
    compare_display = compare_display.rename(columns=rename_map)
    _write(compare_display, TABLE_MAP['Today_vs_Yesterday'])

    # ============================================================
    # BRANCH_TODAY / SOURCE_TODAY
    # ============================================================
    today_branch_full = today_data.groupby(['Branch', 'Segment'], as_index=False).agg(
        Delivered=('Delivered', 'sum'), Workable=('Workable', 'sum'), Prospect=('Prospect', 'sum'),
        Fresh=('Fresh', 'sum'), Junk=('Junk', 'sum'), Current_Adm=('Current Adm', 'sum'),
        Spillover=('Spillover', 'sum'), Total_Adm=('Total Adm', 'sum'),
    )
    today_branch_full['CVR %'] = (today_branch_full['Total_Adm'] / today_branch_full['Delivered'].replace(0, np.nan) * 100).round(2).fillna(0)
    today_branch_full['Junk %'] = (today_branch_full['Junk'] / today_branch_full['Delivered'].replace(0, np.nan) * 100).round(2).fillna(0)
    _write(today_branch_full, TABLE_MAP['Branch_Today'])

    src_today_rows = []
    for src in SOURCES:
        pfx = src.replace(' ', '_') + '_'
        del_c = f"{pfx}Delivered"
        if del_c in today_data.columns:
            row = {
                'Source': src,
                'Delivered': int(today_data[del_c].sum()),
                'Workable': int(today_data.get(f"{pfx}Workable", pd.Series([0])).sum()),
                'Prospect': int(today_data.get(f"{pfx}Prospect", pd.Series([0])).sum()),
                'Fresh': int(today_data.get(f"{pfx}Fresh", pd.Series([0])).sum()),
                'Junk': int(today_data.get(f"{pfx}Junk", pd.Series([0])).sum()),
                'Current Adm': int(today_data.get(f"{pfx}Current_Adm", pd.Series([0])).sum()),
                'Spillover': int(today_data.get(f"{pfx}Spillover", pd.Series([0])).sum()),
                'Total Adm': int(today_data.get(f"{pfx}Total_Adm", pd.Series([0])).sum()),
            }
            row['CVR %'] = round(row['Total Adm'] / max(row['Delivered'], 1) * 100, 2)
            row['Junk %'] = round(row['Junk'] / max(row['Delivered'], 1) * 100, 2)
            src_today_rows.append(row)
    src_today_df = pd.DataFrame(src_today_rows) if src_today_rows else pd.DataFrame(columns=['Source'] + FUNNEL_COLS + ['CVR %', 'Junk %'])
    _write(src_today_df, TABLE_MAP['Source_Today'])

    # ============================================================
    # MONTH-TO-DATE SUMMARY — one row, unambiguous "as of today" view.
    # Every number here is cumulative from the 1st of the month through
    # SNAPSHOT_DATE (not a single day's activity) -- Delivered/Admissions
    # were already computed that way; this just surfaces it as one clean
    # row instead of making you read it off the bottom of a growing ledger.
    # ============================================================
    days_elapsed = (pd.Timestamp(latest_date).date() - pd.Timestamp(CYCLE_START).date()).days + 1

    mtd_row = {
        'As Of Date': today_label,
        'Cycle Start': str(pd.Timestamp(CYCLE_START).date()),
        'Days Elapsed In Cycle': days_elapsed,
        'Total Delivered (MTD)': int(today_data['Delivered'].sum()),
        'Total Workable (MTD)': int(today_data['Workable'].sum()),
        'Total Prospect (MTD)': int(today_data['Prospect'].sum()),
        'Total Fresh (MTD)': int(today_data['Fresh'].sum()),
        'Total Junk (MTD)': int(today_data['Junk'].sum()),
        'Total Admissions (MTD)': int(today_data['Total Adm'].sum()),
    }
    total_del = mtd_row['Total Delivered (MTD)']
    mtd_row['Overall CVR % (MTD)'] = round(mtd_row['Total Admissions (MTD)'] / max(total_del, 1) * 100, 2)
    mtd_row['Overall Junk % (MTD)'] = round(mtd_row['Total Junk (MTD)'] / max(total_del, 1) * 100, 2)

    for src in SOURCES:
        pfx = src.replace(' ', '_') + '_'
        mtd_row[f'{src} Admissions (MTD)'] = int(today_data.get(f'{pfx}Total_Adm', pd.Series([0])).sum())
        mtd_row[f'{src} Delivered (MTD)'] = int(today_data.get(f'{pfx}Delivered', pd.Series([0])).sum())

    month_to_date_summary = pd.DataFrame([mtd_row])
    _write(month_to_date_summary, TABLE_MAP['Month_To_Date_Summary'])

    # ============================================================
    # FULL_DELTA_DETAIL — Branch x University x Segment level
    # ============================================================
    merge_keys = ['Branch', 'University', 'Segment']
    delta = today_data[merge_keys + all_horizontal_metrics].merge(
        yest_data[merge_keys + all_horizontal_metrics] if has_prev else pd.DataFrame(columns=merge_keys + all_horizontal_metrics),
        on=merge_keys, how='outer', suffixes=('_Today', '_Yest'),
    ).fillna(0)
    for m in all_horizontal_metrics:
        t_col, y_col, d_col = f"{m}_Today", f"{m}_Yest", f"{m}_Delta"
        if t_col in delta.columns and y_col in delta.columns:
            delta[d_col] = delta[t_col] - delta[y_col]

    delta['CVR%_Today'] = (delta['Total Adm_Today'] / delta['Delivered_Today'].replace(0, np.nan) * 100).round(2).fillna(0)
    delta['Junk%_Today'] = (delta['Junk_Today'] / delta['Delivered_Today'].replace(0, np.nan) * 100).round(2).fillna(0)

    delta_display = delta[merge_keys + [c for c in delta.columns if c not in merge_keys]].sort_values(['Segment', 'Branch', 'University']).reset_index(drop=True)
    _write(delta_display, TABLE_MAP['Full_Delta_Detail'])

    # ============================================================
    # OVERALL_DASHBOARD (-> daily_overall_dashboard) + TRACK_* TABLES
    # ============================================================
    overall_rows = []
    for d in dates_logged:
        df_d = master[master['Snapshot Date'] == d]
        del_v, wrk_v, pr_v, fr_v, jk_v, adm_v = (
            int(df_d['Delivered'].sum()), int(df_d['Workable'].sum()), int(df_d['Prospect'].sum()),
            int(df_d['Fresh'].sum()), int(df_d['Junk'].sum()), int(df_d['Total Adm'].sum()),
        )
        overall_rows.append({
            'Snapshot Date': str(pd.Timestamp(d).date()), 'Delivered Leads': del_v, 'Workable Leads': wrk_v,
            'Prospect Balance': pr_v, 'Gross New Prospects Gen': pr_v, 'Fresh Leads': fr_v, 'Junk Leads': jk_v,
            'Total Admissions': adm_v, 'CVR %': round(adm_v / max(del_v, 1) * 100, 2), 'Junk %': round(jk_v / max(del_v, 1) * 100, 2),
        })
    overall_dashboard_hist = pd.DataFrame(overall_rows)
    _write(overall_dashboard_hist, TABLE_MAP['Overall_Dashboard'])

    track_table_names = {
        'WhatsApp': 'Track_WhatsApp', 'Live Chat': 'Track_LiveChat', 'AI Meeting': 'Track_AIMeeting',
        'PR': 'Track_PR', 'Website': 'Track_Website', 'RDMPL': 'Track_RDMPL',
    }
    for source_name, info in CHANNEL_BENCHMARKS.items():
        pfx = info['pfx']
        track_rows = []
        for d in dates_logged:
            df_d = master[master['Snapshot Date'] == d]
            s_del = int(df_d.get(f'{pfx}Delivered', pd.Series([0])).sum())
            s_wrk = int(df_d.get(f'{pfx}Workable', pd.Series([0])).sum())
            s_pr = int(df_d.get(f'{pfx}Prospect', pd.Series([0])).sum())
            s_fr = int(df_d.get(f'{pfx}Fresh', pd.Series([0])).sum())
            s_jk = int(df_d.get(f'{pfx}Junk', pd.Series([0])).sum())
            s_adm = int(df_d.get(f'{pfx}Total_Adm', pd.Series([0])).sum())
            track_rows.append({
                'Snapshot Date': str(pd.Timestamp(d).date()), 'Daily Prospect Generation Target': info['p_tgt'],
                'Daily Admissions Milestone Target': info['a_tgt'], 'Delivered Leads': s_del, 'Workable Leads': s_wrk,
                'Prospect Balance': s_pr, 'Gross New Prospects Gen': s_pr, 'Fresh Leads': s_fr, 'Junk Leads': s_jk,
                'Total Admissions': s_adm, 'CVR %': round(s_adm / max(s_del, 1) * 100, 2), 'Junk %': round(s_jk / max(s_del, 1) * 100, 2),
            })
        track_df = pd.DataFrame(track_rows)
        _write(track_df, TABLE_MAP[track_table_names[source_name]])

    return {
        "days_in_history": len(dates_logged),
        "dates_logged": [str(pd.Timestamp(d).date()) for d in dates_logged],
    }


def run_daily_tracker(raw_path: str, enrolled_path: str):
    """Normal daily run: writes ONE row, for SNAPSHOT_DATE (derived from
    the data itself -- max Created On / Modified On), then rebuilds every
    derived table from the full history in Postgres."""
    prep = _prepare_data(raw_path, enrolled_path)
    snap_df = _build_snapshot_df(prep, as_of_date=None)
    _upsert_snapshot(snap_df, prep['SNAPSHOT_DATE'])
    print(f"✅ Step 3 snapshot rows built: {len(snap_df)}")

    rebuild_info = _rebuild_derived_tables()
    print(f"\n🎉 Daily tracker complete for {prep['SNAPSHOT_DATE']}. "
          f"History now spans {rebuild_info['days_in_history']} day(s).")

    return {
        "snapshot_date": str(prep['SNAPSHOT_DATE']),
        "raw_snapshot_date": str(prep['RAW_SNAPSHOT_DATE']),
        "enrolled_snapshot_date": str(prep['ENROLLED_SNAPSHOT_DATE']),
        "days_in_history": rebuild_info['days_in_history'],
        "today_rows": len(snap_df),
        "backfilled": False,
    }


def backfill_daily_history(raw_path: str, enrolled_path: str, from_date=None):
    """
    Reconstructs one row per day from the 1st of the month (or `from_date`,
    as 'YYYY-MM-DD') through SNAPSHOT_DATE, all from this ONE file, using
    each lead's own Created On date as that day's cutoff.

    IMPORTANT LIMITATION: Delivered and Admissions counts are accurate for
    backfilled days (Created On is fixed and can't change after the fact).
    Workable/Prospect/Fresh/Junk are NOT true historical records for
    backfilled days -- they reflect each lead's CURRENT stage today, not
    whatever stage it actually was on that historical day, because the
    source file only has one live snapshot of stage, not a change history.
    Rows from strictly-past days are tagged _backfilled=True in
    daily_snapshots so this is always visible later and never silently
    mixed in with real day-by-day captures. The day equal to SNAPSHOT_DATE
    itself is a true live capture, not an approximation, and is tagged
    _backfilled=False even when run through this function.
    """
    prep = _prepare_data(raw_path, enrolled_path)

    start_date = pd.to_datetime(from_date).date() if from_date else pd.to_datetime(prep['CYCLE_START']).date()
    end_date = prep['SNAPSHOT_DATE']

    if start_date > end_date:
        raise ValueError(f"from_date ({start_date}) is after the snapshot date ({end_date}).")

    all_days = list(pd.date_range(start_date, end_date, freq='D').date)
    print(f"🔁 Backfilling {len(all_days)} day(s): {start_date} -> {end_date}")

    for d in all_days:
        snap_df = _build_snapshot_df(prep, as_of_date=d)
        _upsert_snapshot(snap_df, d)

    rebuild_info = _rebuild_derived_tables()
    print(f"\n🎉 Backfill complete. History now spans {rebuild_info['days_in_history']} day(s).")

    return {
        "snapshot_date": str(prep['SNAPSHOT_DATE']),
        "raw_snapshot_date": str(prep['RAW_SNAPSHOT_DATE']),
        "enrolled_snapshot_date": str(prep['ENROLLED_SNAPSHOT_DATE']),
        "days_in_history": rebuild_info['days_in_history'],
        "backfilled_days": len(all_days),
        "backfilled_range": [str(start_date), str(end_date)],
        "backfilled": True,
    }