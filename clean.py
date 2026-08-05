import streamlit as st
import pandas as pd
import io
import math
import zipfile
from datetime import date, timedelta, datetime
import csv
import gc
import warnings
warnings.filterwarnings("ignore")

# ── CONFIG ───────────────────────────────────────────────────────────────────
PHONE_GROUPS = [
    ("Phone 1", "Phone 1 Type", "Email 1"),
    ("Phone 2", "Phone 2 Type", "Email 2"),
    ("Phone 3", "Phone 3 Type", "Email 3"),
    ("Phone 4", "Phone 4 Type", "Email 4"),
    ("Phone 5", "Phone 5 Type", "Email 5"),
]

DNC_COLS = ["Phone 1 DNC", "Phone 2 DNC", "Phone 3 DNC", "Phone 4 DNC", "Phone 5 DNC"]

MLS_STATUSES = [
    "ACTIVE", "ACTIVE UNDER CONTRACT", "CANCELED", "CANCELLED", "COMING SOON",
    "CONTINGENT", "DELETED", "EXPIRED", "FAIL", "FAILED",
    "PENDING", "REMOVED", "WITHDRAWN"
]

MONTHS = ["JAN","FEB","MAR","APR","MAY","JUN",
          "JUL","AUG","SEP","OCT","NOV","DEC"]

TMP_MERGED = "/tmp/merged.csv"
TMP_PROPS  = "/tmp/props.csv"

# ── HELPERS ──────────────────────────────────────────────────────────────────
def get_val(row, col):
    if col not in row.index:
        return ""
    v = str(row[col]).strip()
    return "" if v.lower() == "nan" else v

def clean_phone(val):
    if not val:
        return ""
    try:
        return str(int(float(val)))
    except:
        return val

def build_campaign(channel, tail, tag=None, wk=None):
    tag_part = f" - {tag}" if tag else ""
    wk_part  = f" - WK{wk}" if wk else ""
    return f"{channel}{tag_part}{wk_part} - {tail}"

def calc_week_ranges(total, n_splits):
    chunk = math.ceil(total / n_splits)
    ranges = []
    for i in range(n_splits):
        start = i * chunk + 1
        end   = min((i + 1) * chunk, total)
        ranges.append(f"WK{i+1}: {start:,} – {end:,} ({end-start+1:,} contacts)")
    return ranges

# ── MERGE TO DISK (memory safe) ──────────────────────────────────────────────
def merge_to_disk(files, tmp_path):
    """Reads files one by one and writes to disk. Returns (total_rows, orig_phones, orig_emails)."""
    total_rows  = 0
    orig_phones = 0
    orig_emails = 0
    phone_cols  = [p for p, t, e in PHONE_GROUPS]
    email_cols  = [e for p, t, e in PHONE_GROUPS]
    header_written = False

    with open(tmp_path, "w", newline="", encoding="utf-8") as out_f:
        for uf in files:
            df = pd.read_excel(uf, dtype=str)
            df.drop(columns=[c for c in DNC_COLS if c in df.columns], inplace=True)
            for col in phone_cols:
                if col in df.columns:
                    orig_phones += df[col].apply(
                        lambda x: 1 if pd.notna(x) and str(x).strip() not in ("","nan") else 0).sum()
            for col in email_cols:
                if col in df.columns:
                    orig_emails += df[col].apply(
                        lambda x: 1 if pd.notna(x) and str(x).strip() not in ("","nan") else 0).sum()
            df.to_csv(out_f, index=False, header=not header_written)
            total_rows += len(df)
            header_written = True
            del df
            gc.collect()

    return total_rows, int(orig_phones), int(orig_emails)

# ── FILTER PROPERTIES TO DISK ────────────────────────────────────────────────
def filter_properties_to_disk(src_path, dst_path):
    df = pd.read_csv(src_path, dtype=str)
    if "MLS Status" not in df.columns:
        pd.DataFrame(columns=df.columns).to_csv(dst_path, index=False)
        return 0
    blank_mask = df["MLS Status"].apply(lambda x: str(x).strip() in ("", "nan"))
    known_mask = df["MLS Status"].apply(lambda x: str(x).strip().upper() in MLS_STATUSES)
    known  = df[known_mask].copy()
    blanks = df[blank_mask].copy()
    blanks["MLS Status"] = "Off Market"
    result = pd.concat([known, blanks], ignore_index=True)
    result.to_csv(dst_path, index=False)
    rows = len(result)
    del df, known, blanks, result
    gc.collect()
    return rows

# ── PROCESS FUNCTIONS (disk → disk) ─────────────────────────────────────────
def process_dialer_csv(src, campaign_name):
    df = pd.read_csv(src, dtype=str)
    all_phone_cols = [p for p, t, e in PHONE_GROUPS] + [t for p, t, e in PHONE_GROUPS]
    other_cols     = [c for c in df.columns if c not in all_phone_cols]
    active_groups  = [(p, t) for p, t, e in PHONE_GROUPS if p in df.columns]
    output_cols    = ["Campaign Name", "Phone Number", "Phone Type"] + other_cols
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=output_cols)
    writer.writeheader()
    total = 0
    for _, row in df.iterrows():
        for phone_col, type_col in active_groups:
            phone_val = clean_phone(get_val(row, phone_col))
            if not phone_val:
                continue
            new_row = {"Campaign Name": campaign_name, "Phone Number": phone_val,
                       "Phone Type": get_val(row, type_col)}
            for col in other_cols:
                new_row[col] = get_val(row, col)
            writer.writerow(new_row)
            total += 1
    del df
    gc.collect()
    return total, buf.getvalue().encode("utf-8")

def process_sms_csv(src, campaign_name):
    df = pd.read_csv(src, dtype=str)
    all_pe = ([p for p, t, e in PHONE_GROUPS] + [t for p, t, e in PHONE_GROUPS] +
              [e for p, t, e in PHONE_GROUPS if e in df.columns])
    other_cols    = [c for c in df.columns if c not in all_pe]
    active_groups = [(p, t, e) for p, t, e in PHONE_GROUPS if p in df.columns or e in df.columns]
    output_cols   = ["Campaign Name", "Phone Number", "Phone Type", "Email"] + other_cols
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=output_cols)
    writer.writeheader()
    total = 0
    for _, row in df.iterrows():
        for phone_col, type_col, email_col in active_groups:
            phone_val  = clean_phone(get_val(row, phone_col))
            phone_type = get_val(row, type_col)
            email_val  = get_val(row, email_col)
            if phone_type.strip().lower() == "landline":
                phone_val = ""
            if not phone_val and not email_val:
                continue
            new_row = {"Campaign Name": campaign_name, "Phone Number": phone_val,
                       "Phone Type": phone_type, "Email": email_val}
            for col in other_cols:
                new_row[col] = get_val(row, col)
            writer.writerow(new_row)
            total += 1
    del df
    gc.collect()
    return total, buf.getvalue().encode("utf-8")

def process_email_csv(src, campaign_name):
    df = pd.read_csv(src, dtype=str)
    all_pe = ([p for p, t, e in PHONE_GROUPS] + [t for p, t, e in PHONE_GROUPS] +
              [e for p, t, e in PHONE_GROUPS if e in df.columns])
    other_cols    = [c for c in df.columns if c not in all_pe]
    active_groups = [(p, t, e) for p, t, e in PHONE_GROUPS if p in df.columns or e in df.columns]
    output_cols   = ["Campaign Name", "Phone Number", "Phone Type", "Email"] + other_cols
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=output_cols)
    writer.writeheader()
    total = 0
    for _, row in df.iterrows():
        for phone_col, type_col, email_col in active_groups:
            email_val = get_val(row, email_col)
            if not email_val:
                continue
            phone_val  = clean_phone(get_val(row, phone_col))
            phone_type = get_val(row, type_col)
            if phone_type.strip().lower() == "landline":
                phone_val = ""
            new_row = {"Campaign Name": campaign_name, "Phone Number": phone_val,
                       "Phone Type": phone_type, "Email": email_val}
            for col in other_cols:
                new_row[col] = get_val(row, col)
            writer.writerow(new_row)
            total += 1
    del df
    gc.collect()
    return total, buf.getvalue().encode("utf-8")

# ── SPLIT CSV BYTES ──────────────────────────────────────────────────────────
def split_csv_bytes(csv_bytes, name_builder, n_splits, do_split):
    df = pd.read_csv(io.StringIO(csv_bytes.decode("utf-8")), dtype=str)
    if df.empty:
        return []
    if do_split and n_splits > 1:
        total = len(df)
        chunk = math.ceil(total / n_splits)
        parts = []
        for i in range(n_splits):
            c = df.iloc[i * chunk:(i + 1) * chunk].copy()
            if c.empty:
                continue
            name = name_builder(i + 1)
            c["Campaign Name"] = name
            b = io.StringIO()
            c.to_csv(b, index=False)
            parts.append((name, b.getvalue().encode("utf-8")))
        return parts
    else:
        name = name_builder(None)
        return [(name, csv_bytes)]

# ── BUILD ZIP ────────────────────────────────────────────────────────────────
def build_zip(channel_parts):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for folder, parts in channel_parts.items():
            for name, data in parts:
                zf.writestr(f"{folder}/{name}.csv", data)
    buf.seek(0)
    return buf.getvalue()

# ── RUN CHANNEL PROCESSING ───────────────────────────────────────────────────
def run_channels(src_path, tail, tag, do_dialer, do_sms, do_email, do_split, n_splits):
    channel_parts = {}
    counts = {}
    if do_dialer:
        campaign = build_campaign("CC", tail, tag=tag)
        rows, data = process_dialer_csv(src_path, campaign)
        counts["dialer"] = rows
        parts = split_csv_bytes(data, lambda wk: build_campaign("CC", tail, tag=tag, wk=wk), n_splits, do_split)
        channel_parts["CC"] = parts
    if do_sms:
        campaign = build_campaign("SMS", tail, tag=tag)
        rows, data = process_sms_csv(src_path, campaign)
        counts["sms"] = rows
        parts = split_csv_bytes(data, lambda wk: build_campaign("SMS", tail, tag=tag, wk=wk), n_splits, do_split)
        channel_parts["SMS"] = parts
    if do_email:
        campaign = build_campaign("EMAIL", tail, tag=tag)
        rows, data = process_email_csv(src_path, campaign)
        counts["email"] = rows
        parts = split_csv_bytes(data, lambda wk: build_campaign("EMAIL", tail, tag=tag, wk=wk), n_splits, do_split)
        channel_parts["EMAIL"] = parts
    return channel_parts, counts

# ── UI HELPERS ───────────────────────────────────────────────────────────────
def step_header(num, icon, title, optional=False):
    col1, col2 = st.columns([5, 1])
    with col1:
        st.markdown(f"##### {icon}&nbsp;&nbsp;**Step {num} — {title}**")
    with col2:
        if optional:
            st.badge("Optional", color="orange")
    st.markdown("---")

def page_title(icon, title, subtitle):
    with st.container(border=True):
        col1, col2 = st.columns([1, 8])
        with col1:
            st.markdown(f"# {icon}")
        with col2:
            st.markdown(f"### {title}")
            st.caption(subtitle)

def campaign_details_block(key_prefix):
    add_name = st.checkbox("Do you want to add Campaign Name?", key=f"{key_prefix}_toggle")
    month = year = state = deal = output_name = None
    if add_name:
        col1, col2 = st.columns(2)
        with col1:
            month = st.selectbox("📅 Month", MONTHS, key=f"{key_prefix}_month")
            state = st.text_input("🗺️ State", placeholder="e.g. FL", key=f"{key_prefix}_state")
        with col2:
            year = st.text_input("📆 Year", value="2026", key=f"{key_prefix}_year")
            deal = st.text_input("🏷️ Type of Deal", placeholder="e.g. LUXURY LAND", key=f"{key_prefix}_deal")
    else:
        output_name = st.text_input("Please enter the Output File Name", key=f"{key_prefix}_outname")
    return add_name, month, year, state, deal, output_name

def get_tail(add_name, month, year, state, deal, output_name):
    if add_name:
        return f"{month} - {year} - {state.strip().upper()} - {deal.strip()}"
    return output_name.strip() if output_name else ""

def marketing_process_block(key_prefix):
    col1, col2, col3 = st.columns(3)
    with col1:
        dialer = st.checkbox("📞 Dialer / AI Outbound", key=f"{key_prefix}_dialer")
    with col2:
        sms = st.checkbox("💬 SMS", key=f"{key_prefix}_sms")
    with col3:
        email = st.checkbox("📧 Email", key=f"{key_prefix}_email")
    return dialer, sms, email

# ── PAGE: PEOPLE LEADS ───────────────────────────────────────────────────────
def page_people_leads():
    page_title("👤", "People Leads", "Upload, clean, and organize your people lead campaigns")

    step_header(1, "📁", "Upload Lead List")
    files = st.file_uploader("Upload one or more Excel files (.xlsx)",
                              type=["xlsx"], accept_multiple_files=True, key="ppl_upload",
                              label_visibility="collapsed")
    if files:
        st.success(f"✅ {len(files)} file(s) uploaded")

    step_header(2, "🏷️", "Campaign Details")
    add_name, month, year, state, deal, out_name = campaign_details_block("ppl")

    step_header(3, "📣", "Marketing Process")
    dialer, sms, email = marketing_process_block("ppl_mkt")

    step_header(4, "🏠", "Properties / Seller Leads", optional=True)
    p_dialer, p_sms, p_email = marketing_process_block("ppl_prop")

    step_header(5, "🔀", "Monthly Split", optional=True)
    do_split = st.checkbox("🔀 Split output into multiple files", key="ppl_dosplit")
    n_splits = 1
    if do_split:
        n_splits = st.number_input("How many files to split into?", min_value=2, max_value=50,
                                   value=5, step=1, key="ppl_nsplits")

    mkt_selected  = dialer or sms or email
    prop_selected = p_dialer or p_sms or p_email

    if add_name:
        details_ok = bool(state and state.strip() and year and year.strip() and deal and deal.strip())
    else:
        details_ok = bool(out_name and out_name.strip())

    ready = bool(files and details_ok and (mkt_selected or prop_selected))

    if st.button("⚙️ Process Files", width='stretch', type="primary",
                 disabled=not ready, key="ppl_process"):
        try:
            tail = get_tail(add_name, month, year, state, deal, out_name)

            with st.spinner("📥 Merging files... please wait"):
                total_rows, orig_phones, orig_emails = merge_to_disk(files, TMP_MERGED)
            st.info(f"📊 Merged {total_rows:,} rows from {len(files)} file(s)")

            mkt_parts  = {}
            mkt_counts = {}
            if mkt_selected:
                with st.spinner("⚙️ Processing Marketing channels..."):
                    mkt_parts, mkt_counts = run_channels(
                        TMP_MERGED, tail, None, dialer, sms, email, do_split, int(n_splits))

            prop_parts  = {}
            prop_counts = {}
            if prop_selected:
                with st.spinner("🏠 Filtering Properties / Seller Leads..."):
                    filter_properties_to_disk(TMP_MERGED, TMP_PROPS)
                with st.spinner("⚙️ Processing Properties channels..."):
                    prop_parts, prop_counts = run_channels(
                        TMP_PROPS, tail, "PROPERTIES", p_dialer, p_sms, p_email, do_split, int(n_splits))

            st.session_state.ppl_mkt_zip      = build_zip(mkt_parts) if mkt_parts else None
            st.session_state.ppl_mkt_name     = f"Marketing Process - {tail}.zip"
            st.session_state.ppl_mkt_counts   = mkt_counts
            st.session_state.ppl_prop_zip     = build_zip(prop_parts) if prop_parts else None
            st.session_state.ppl_prop_name    = f"Properties Seller Leads - {tail}.zip"
            st.session_state.ppl_prop_counts  = prop_counts
            st.session_state.ppl_total        = total_rows
            st.session_state.ppl_orig_phones  = orig_phones
            st.session_state.ppl_orig_emails  = orig_emails
            st.session_state.ppl_processed    = True
            gc.collect()

        except Exception as e:
            st.error(f"❌ Error: {e}")

    if st.session_state.get("ppl_processed"):
        st.success("✅ **Processing Complete!**")
        st.info(f"📊 Total rows merged: **{st.session_state.ppl_total:,}**")
        st.info(f"📞 Original phone numbers: **{st.session_state.ppl_orig_phones:,}**")
        st.info(f"📧 Original emails: **{st.session_state.ppl_orig_emails:,}**")

        if st.session_state.get("ppl_mkt_counts"):
            st.markdown("**📋 Marketing Process:**")
            for k, v in st.session_state.ppl_mkt_counts.items():
                st.success(f"{k.title()}: **{v:,}**")
        if st.session_state.get("ppl_prop_counts"):
            st.markdown("**🏠 Properties / Seller Leads:**")
            for k, v in st.session_state.ppl_prop_counts.items():
                st.success(f"{k.title()}: **{v:,}**")

        step_header("⬇", "📦", "Downloads")
        if st.session_state.get("ppl_mkt_zip"):
            st.download_button("⬇️ Download Marketing Process (ZIP)",
                               data=st.session_state.ppl_mkt_zip,
                               file_name=st.session_state.ppl_mkt_name,
                               mime="application/zip", width='stretch',
                               type="primary", key="ppl_dl_mkt")
        if st.session_state.get("ppl_prop_zip"):
            st.download_button("⬇️ Download Properties / Seller Leads (ZIP)",
                               data=st.session_state.ppl_prop_zip,
                               file_name=st.session_state.ppl_prop_name,
                               mime="application/zip", width='stretch',
                               type="primary", key="ppl_dl_prop")

    if not ready and files:
        if add_name:
            if not (state and state.strip()):
                st.warning("⚠️ Please enter State.")
            if not (year and year.strip()):
                st.warning("⚠️ Please enter Year.")
            if not (deal and deal.strip()):
                st.warning("⚠️ Please enter Type of Deal.")
        else:
            if not (out_name and out_name.strip()):
                st.warning("⚠️ Please enter the Output File Name.")
        if not (mkt_selected or prop_selected):
            st.warning("⚠️ Please select at least one output type.")

# ── PAGE: BUSINESS LEADS ─────────────────────────────────────────────────────
def page_business_leads():
    page_title("🏢", "Business Leads", "Merge your business lead files and export in one click")

    step_header(1, "📁", "Upload Lead List")
    files = st.file_uploader("Upload one or more Excel or CSV files (.xlsx, .csv)",
                              type=["xlsx","csv"], accept_multiple_files=True, key="biz_upload",
                              label_visibility="collapsed")
    if files:
        st.success(f"✅ {len(files)} file(s) uploaded")

    step_header(2, "🏷️", "Campaign Name")
    add_name, month, year, state, deal, out_name = campaign_details_block("biz")

    step_header(3, "🔀", "Monthly Split", optional=True)
    do_split = st.checkbox("🔀 Split output into multiple files", key="biz_dosplit")
    n_splits = 1
    if do_split:
        n_splits = st.number_input("How many files to split into?", min_value=2, max_value=50,
                                   value=5, step=1, key="biz_nsplits")

    if add_name:
        details_ok = bool(state and state.strip() and year and year.strip() and deal and deal.strip())
    else:
        details_ok = bool(out_name and out_name.strip())

    ready = bool(files and details_ok)

    if st.button("⚙️ Process Files", width='stretch', type="primary",
                 disabled=not ready, key="biz_process"):
        try:
            tail = get_tail(add_name, month, year, state, deal, out_name)

            with st.spinner("📥 Merging files..."):
                frames = []
                for uf in files:
                    if uf.name.lower().endswith(".csv"):
                        frames.append(pd.read_csv(uf, dtype=str))
                    else:
                        frames.append(pd.read_excel(uf, dtype=str))
                merged = pd.concat(frames, ignore_index=True)
                del frames
                gc.collect()

            row_count = len(merged)
            name_builder = lambda wk: build_campaign("LEADS", tail, wk=wk)
            base_parts = []
            if do_split and n_splits > 1:
                chunk = math.ceil(row_count / int(n_splits))
                for i in range(int(n_splits)):
                    c = merged.iloc[i*chunk:(i+1)*chunk].copy()
                    if c.empty:
                        continue
                    name = name_builder(i+1)
                    c["Campaign Name"] = name
                    b = io.StringIO()
                    c.to_csv(b, index=False)
                    base_parts.append((name, b.getvalue().encode("utf-8")))
            else:
                b = io.StringIO()
                merged.to_csv(b, index=False)
                base_parts.append((name_builder(None), b.getvalue().encode("utf-8")))

            del merged
            gc.collect()

            st.session_state.biz_base_zip   = build_zip({"LEADS": base_parts})
            st.session_state.biz_base_name  = f"Business Leads - {tail}.zip"
            st.session_state.biz_total      = row_count
            st.session_state.biz_processed  = True

        except Exception as e:
            st.error(f"❌ Error: {e}")

    if st.session_state.get("biz_processed"):
        st.success("✅ **Processing Complete!**")
        st.info(f"📊 Total rows merged: **{st.session_state.biz_total:,}**")
        step_header("⬇", "📦", "Downloads")
        if st.session_state.get("biz_base_zip"):
            st.download_button("⬇️ Download Business Leads (ZIP)",
                               data=st.session_state.biz_base_zip,
                               file_name=st.session_state.biz_base_name,
                               mime="application/zip", width='stretch',
                               type="primary", key="biz_dl")

    if not ready and files:
        if add_name:
            if not (state and state.strip()): st.warning("⚠️ Please enter State.")
            if not (year and year.strip()):   st.warning("⚠️ Please enter Year.")
            if not (deal and deal.strip()):   st.warning("⚠️ Please enter Type of Deal.")
        else:
            if not (out_name and out_name.strip()): st.warning("⚠️ Please enter Output File Name.")

# ── REPORTING CONSTANTS ───────────────────────────────────────────────────────
MAX_DIALS_PER_NUMBER    = 150
DIALS_PER_CALLER        = 600
MIN_CALLER_ID_POOL      = 10
ROTATION_CALLING_DAYS   = 10
DIALS_BEFORE_ROTATION   = 1500
CALLING_DAYS_PER_MONTH  = 20

MAX_SMS_PER_NUMBER      = 500
SMS_NUMBERS_PER_CAMPAIGN= 4
MAX_SMS_PER_CAMPAIGN    = 2000
SMS_ROTATION_DAYS       = 10
SMS_BEFORE_ROTATION     = 5000

MAX_INBOXES_PER_DOMAIN  = 5
MAX_EMAILS_PER_INBOX    = 35
DOMAIN_DAILY_CAPACITY   = 175
EMAIL_WARMUP_DAYS       = 30
EMAIL_ROTATION_DAYS     = 20
EMAILS_BEFORE_ROTATION  = 700

# ── ACCENT PALETTE (light theme, indigo-based) ────────────────────────────────
CLR_INDIGO = "#3346D3"
CLR_BLUE   = "#3B82F6"
CLR_GREEN  = "#22A45D"
CLR_PURPLE = "#8B5CF6"
CLR_RED    = "#EF4444"
CLR_AMBER  = "#F59E0B"
CLR_INK    = "#1D2140"
CLR_MUTE   = "#6B7290"
CLR_GRID   = "#E7EAF3"

def add_business_days(start_date, days):
    current = start_date
    added = 0
    while added < days:
        current += timedelta(days=1)
        if current.weekday() < 5:
            added += 1
    return current

def calculate_all(total_phones, total_sms, total_emails, callers, start_date=None):
    r = {}
    r["dials_per_day"]        = math.ceil(total_phones / CALLING_DAYS_PER_MONTH)
    r["callers_needed"]       = math.ceil(r["dials_per_day"] / DIALS_PER_CALLER)
    r["active_caller_ids"]    = max(MIN_CALLER_ID_POOL, math.ceil(r["dials_per_day"] / MAX_DIALS_PER_NUMBER))
    r["rotation_cycles"]      = CALLING_DAYS_PER_MONTH // ROTATION_CALLING_DAYS
    r["total_caller_ids"]     = r["active_caller_ids"] * r["rotation_cycles"]
    r["dials_per_number"]     = math.ceil(r["dials_per_day"] / r["active_caller_ids"])
    r["dials_check_ok"]       = r["dials_per_number"] <= MAX_DIALS_PER_NUMBER

    r["callers_diff"]         = callers - r["callers_needed"]
    r["staffing_ok"]          = callers >= r["callers_needed"]
    r["team_capacity"]        = callers * DIALS_PER_CALLER * CALLING_DAYS_PER_MONTH
    r["days_to_finish"]       = math.ceil(total_phones / (callers * DIALS_PER_CALLER)) if callers > 0 else 999

    r["texts_per_day"]        = math.ceil(total_sms / CALLING_DAYS_PER_MONTH)
    r["min_sms_numbers"]      = math.ceil(r["texts_per_day"] / MAX_SMS_PER_NUMBER)
    r["campaigns_required"]   = math.ceil(r["texts_per_day"] / MAX_SMS_PER_CAMPAIGN)
    r["active_sms_numbers"]   = max(r["min_sms_numbers"], r["campaigns_required"] * SMS_NUMBERS_PER_CAMPAIGN)
    r["total_sms_numbers"]    = r["active_sms_numbers"] * r["rotation_cycles"]
    r["texts_per_number"]     = math.ceil(r["texts_per_day"] / r["active_sms_numbers"]) if r["active_sms_numbers"] > 0 else 0
    r["sms_check_ok"]         = r["texts_per_number"] <= MAX_SMS_PER_NUMBER

    r["emails_per_day"]       = math.ceil(total_emails / CALLING_DAYS_PER_MONTH)
    r["inboxes_needed"]       = math.ceil(r["emails_per_day"] / MAX_EMAILS_PER_INBOX)
    r["domains_needed"]       = math.ceil(r["inboxes_needed"] / MAX_INBOXES_PER_DOMAIN)
    r["domain_capacity"]      = r["domains_needed"] * DOMAIN_DAILY_CAPACITY
    r["emails_per_inbox"]     = math.ceil(r["emails_per_day"] / r["inboxes_needed"]) if r["inboxes_needed"] > 0 else 0
    r["email_check_ok"]       = r["emails_per_inbox"] <= MAX_EMAILS_PER_INBOX

    if start_date:
        r["campaign_end"]       = add_business_days(start_date, CALLING_DAYS_PER_MONTH)
        r["rotate_out_1"]       = add_business_days(start_date, ROTATION_CALLING_DAYS)
        r["rotate_out_2"]       = add_business_days(start_date, ROTATION_CALLING_DAYS * 2)
        r["warmup_start"]       = start_date - timedelta(days=EMAIL_WARMUP_DAYS)
        r["email_rotate"]       = add_business_days(start_date, EMAIL_ROTATION_DAYS)
    else:
        for k in ("campaign_end","rotate_out_1","rotate_out_2","warmup_start","email_rotate"):
            r[k] = None
    return r

# ── PLOTLY CHART BUILDERS ─────────────────────────────────────────────────────
def _base_layout(fig, height=260):
    fig.update_layout(
        height=height,
        margin=dict(l=10, r=10, t=30, b=10),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter, Segoe UI, sans-serif", color=CLR_INK, size=13),
        showlegend=False,
    )
    return fig

def gauge_chart(value, limit, title, unit=""):
    import plotly.graph_objects as go
    ok = value <= limit
    bar_color = CLR_GREEN if ok else CLR_RED
    axis_max = max(limit * 1.4, value * 1.15, 1)
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=value,
        number={"suffix": unit, "font": {"size": 30, "color": CLR_INK}},
        title={"text": title, "font": {"size": 14, "color": CLR_MUTE}},
        gauge={
            "axis": {"range": [0, axis_max], "tickcolor": CLR_MUTE, "tickfont": {"size": 10}},
            "bar": {"color": bar_color, "thickness": 0.7},
            "bgcolor": "#F4F6FB",
            "borderwidth": 0,
            "steps": [
                {"range": [0, limit], "color": "#E8F5EC"},
                {"range": [limit, axis_max], "color": "#FCE9E9"},
            ],
            "threshold": {"line": {"color": CLR_INK, "width": 3}, "thickness": 0.8, "value": limit},
        },
    ))
    return _base_layout(fig, height=240)

def bar_breakdown(labels, values, colors, title):
    import plotly.graph_objects as go
    fig = go.Figure(go.Bar(
        x=labels, y=values,
        marker=dict(color=colors, line=dict(width=0)),
        text=[f"{v:,}" for v in values],
        textposition="outside",
        textfont=dict(size=12, color=CLR_INK),
    ))
    fig.update_yaxes(showgrid=True, gridcolor=CLR_GRID, zeroline=False)
    fig.update_xaxes(showgrid=False)
    fig.update_layout(title=dict(text=title, font=dict(size=14, color=CLR_MUTE)))
    return _base_layout(fig, height=300)

def timeline_chart(events):
    """events: list of dicts {task, start, end, color}"""
    import plotly.graph_objects as go
    fig = go.Figure()
    for i, ev in enumerate(events):
        fig.add_trace(go.Bar(
            base=[ev["start"]],
            x=[(ev["end"] - ev["start"]).days],
            y=[ev["task"]],
            orientation="h",
            marker=dict(color=ev["color"]),
            hovertemplate=f"{ev['task']}<br>{ev['start'].strftime('%b %d')} → {ev['end'].strftime('%b %d')}<extra></extra>",
            width=0.55,
        ))
    fig.update_xaxes(type="date", showgrid=True, gridcolor=CLR_GRID)
    fig.update_yaxes(showgrid=False, autorange="reversed")
    fig.update_layout(barmode="stack")
    return _base_layout(fig, height=260)

# ── LIGHT KPI CARD (native, no dark HTML) ─────────────────────────────────────
def kpi_card(col, label, value, desc, status=None):
    """status: None | 'ok' | 'bad' -> tints the value color."""
    color = CLR_INK
    if status == "ok":
        color = CLR_GREEN
    elif status == "bad":
        color = CLR_RED
    with col:
        st.markdown(
            f"""<div style="background:#FFFFFF;border:1px solid {CLR_GRID};border-radius:14px;
                        padding:16px 18px;margin-bottom:12px;box-shadow:0 1px 3px rgba(20,25,45,0.04);">
                    <div style="font-size:0.72rem;font-weight:600;color:{CLR_MUTE};
                                text-transform:uppercase;letter-spacing:1px;margin-bottom:6px;">{label}</div>
                    <div style="font-size:1.9rem;font-weight:800;color:{color};line-height:1;">{value}</div>
                    <div style="font-size:0.72rem;color:#9AA0B4;margin-top:5px;">{desc}</div>
                </div>""",
            unsafe_allow_html=True,
        )

def channel_banner(icon, title, subtitle, accent):
    st.markdown(
        f"""<div style="display:flex;align-items:center;gap:12px;background:#FFFFFF;
                    border:1px solid {CLR_GRID};border-left:5px solid {accent};
                    border-radius:12px;padding:14px 18px;margin:8px 0 16px 0;">
                <span style="font-size:1.7rem;">{icon}</span>
                <div>
                    <div style="font-size:1.05rem;font-weight:800;color:{CLR_INK};">{title}</div>
                    <div style="font-size:0.75rem;color:{CLR_MUTE};">{subtitle}</div>
                </div>
            </div>""",
        unsafe_allow_html=True,
    )

# ── REPORTING PAGE ────────────────────────────────────────────────────────────
def render_reporting(auto_phones=0, auto_sms=0, auto_emails=0):
    page_title("📊", "Marketing Guardrails", "Monthly campaign capacity & compliance dashboard")

    st.markdown("##### 🎯&nbsp;&nbsp;**Campaign Inputs**")
    with st.container(border=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            total_phones = st.number_input("📞 Total Phone Numbers", min_value=0,
                value=int(auto_phones) if auto_phones else 0, step=1000, format="%d", key="rpt_phones")
        with c2:
            total_sms = st.number_input("💬 Total SMS Numbers (Mobile)", min_value=0,
                value=int(auto_sms) if auto_sms else 0, step=1000, format="%d", key="rpt_sms")
        with c3:
            total_emails = st.number_input("📧 Total Email Addresses", min_value=0,
                value=int(auto_emails) if auto_emails else 0, step=1000, format="%d", key="rpt_emails")
        c4, c5 = st.columns(2)
        with c4:
            callers = st.number_input("👥 Callers on Team", min_value=0, value=2, step=1, format="%d", key="rpt_callers")
        with c5:
            use_dates = st.checkbox("📅 Add Campaign Dates (Optional)", key="rpt_use_dates")
        start_date = None
        if use_dates:
            start_date = st.date_input("Campaign Start Date", value=date.today(), key="rpt_start")

    if total_phones == 0 and total_sms == 0 and total_emails == 0:
        st.info("👆 Enter your numbers above to generate the report. Process your files first and the counts auto-populate here.")
        return

    r = calculate_all(total_phones, total_sms, total_emails, callers, start_date)

    # ── OVERALL COMPLIANCE STATUS ─────────────────────────────────────────────
    checks = [("Cold Calling", r["dials_check_ok"]), ("SMS", r["sms_check_ok"]),
              ("Email", r["email_check_ok"]), ("Staffing", r["staffing_ok"])]
    passed = sum(1 for _, ok in checks if ok)
    st.markdown("")
    if passed == len(checks):
        st.success(f"✅ **All {len(checks)} guardrails within limits** — campaign is safe to launch.")
    else:
        failed = [name for name, ok in checks if not ok]
        st.warning(f"⚠️ **{passed}/{len(checks)} guardrails OK** — needs attention: {', '.join(failed)}")

    st.divider()

    # ══ COLD CALLING ══════════════════════════════════════════════════════════
    channel_banner("📞", "Cold Calling", "Landline + Mobile · all phone numbers", CLR_BLUE)
    k = st.columns(4)
    kpi_card(k[0], "Dials / Day", f"{r['dials_per_day']:,}", "Total ÷ 20 days")
    kpi_card(k[1], "Callers Needed", f"{r['callers_needed']:,}", "Dials ÷ 600/caller")
    kpi_card(k[2], "Active Caller IDs", f"{r['active_caller_ids']:,}", "Per rotation cycle")
    kpi_card(k[3], "Total Caller IDs", f"{r['total_caller_ids']:,}", f"Active × {r['rotation_cycles']} cycles")

    g1, g2 = st.columns([1, 1])
    with g1:
        st.plotly_chart(gauge_chart(r["dials_per_number"], MAX_DIALS_PER_NUMBER,
                        "Dials / Number / Day", ""), width='stretch',
                        config={"displayModeBar": False}, key="g_cc")
        st.caption(f"Limit: {MAX_DIALS_PER_NUMBER}/number/day · "
                   + ("✅ within limit" if r['dials_check_ok'] else "⛔ over limit"))
    with g2:
        st.plotly_chart(bar_breakdown(
            ["Dials/Day", "Team Cap/Day", "Per Caller"],
            [r["dials_per_day"], callers * DIALS_PER_CALLER, DIALS_PER_CALLER],
            [CLR_BLUE, CLR_INDIGO, "#9AA0E8"],
            "Daily Calling Capacity"), width='stretch',
            config={"displayModeBar": False}, key="b_cc")

    st.divider()

    # ══ STAFFING ══════════════════════════════════════════════════════════════
    st.markdown("##### 👥&nbsp;&nbsp;**Staffing Analysis**")
    if r["staffing_ok"]:
        st.success(f"✅ **Sufficient** — {callers} callers, {r['callers_diff']} surplus over the "
                   f"{r['callers_needed']} needed. List finishes in ~{r['days_to_finish']} days.")
    else:
        st.error(f"⚠️ **Shortfall** — {callers} callers but need {r['callers_needed']}. "
                 f"Hire {abs(r['callers_diff'])} more to finish in 20 days "
                 f"(currently ~{r['days_to_finish']} days).")
    sc = st.columns(3)
    kpi_card(sc[0], "Callers Needed", f"{r['callers_needed']:,}", "To finish in 20 days")
    kpi_card(sc[1], "Team Capacity (20d)", f"{r['team_capacity']:,}", f"{callers} × 600 × 20")
    kpi_card(sc[2], "Days to Finish", f"{r['days_to_finish']}", "Target: ≤ 20",
             status="ok" if r["days_to_finish"] <= 20 else "bad")

    st.divider()

    # ══ SMS ═══════════════════════════════════════════════════════════════════
    channel_banner("💬", "SMS", "Mobile numbers only · A2P 10DLC", CLR_GREEN)
    sk = st.columns(4)
    kpi_card(sk[0], "Texts / Day", f"{r['texts_per_day']:,}", "SMS ÷ 20 days")
    kpi_card(sk[1], "Campaigns Required", f"{r['campaigns_required']:,}", "÷ 2,000/campaign/day")
    kpi_card(sk[2], "Active SMS Numbers", f"{r['active_sms_numbers']:,}", f"{SMS_NUMBERS_PER_CAMPAIGN}/campaign")
    kpi_card(sk[3], "Total SMS Numbers", f"{r['total_sms_numbers']:,}", f"Active × {r['rotation_cycles']} cycles")

    sg1, sg2 = st.columns([1, 1])
    with sg1:
        st.plotly_chart(gauge_chart(r["texts_per_number"], MAX_SMS_PER_NUMBER,
                        "Texts / Number / Day", ""), width='stretch',
                        config={"displayModeBar": False}, key="g_sms")
        st.caption(f"Limit: {MAX_SMS_PER_NUMBER}/number/day · "
                   + ("✅ within limit" if r['sms_check_ok'] else "⛔ over limit"))
    with sg2:
        st.plotly_chart(bar_breakdown(
            ["Texts/Day", "Min Numbers", "Active Numbers"],
            [r["texts_per_day"], r["min_sms_numbers"], r["active_sms_numbers"]],
            [CLR_GREEN, "#7DD3A8", CLR_INDIGO],
            "SMS Volume & Numbers"), width='stretch',
            config={"displayModeBar": False}, key="b_sms")

    st.divider()

    # ══ EMAIL ═════════════════════════════════════════════════════════════════
    channel_banner("📧", "Email", "Verified emails only · CAN-SPAM compliant", CLR_PURPLE)
    ek = st.columns(4)
    kpi_card(ek[0], "Emails / Day", f"{r['emails_per_day']:,}", "Total ÷ 20 days")
    kpi_card(ek[1], "Inboxes Needed", f"{r['inboxes_needed']:,}", "÷ 35/inbox/day")
    kpi_card(ek[2], "Domains Needed", f"{r['domains_needed']:,}", "÷ 5 inboxes/domain")
    kpi_card(ek[3], "Domain Daily Cap", f"{r['domain_capacity']:,}", "Domains × 175/day")

    eg1, eg2 = st.columns([1, 1])
    with eg1:
        st.plotly_chart(gauge_chart(r["emails_per_inbox"], MAX_EMAILS_PER_INBOX,
                        "Emails / Inbox / Day", ""), width='stretch',
                        config={"displayModeBar": False}, key="g_email")
        st.caption(f"Limit: {MAX_EMAILS_PER_INBOX}/inbox/day · "
                   + ("✅ within limit" if r['email_check_ok'] else "⛔ over limit"))
    with eg2:
        st.plotly_chart(bar_breakdown(
            ["Inboxes", "Domains", "Cap ÷ 100"],
            [r["inboxes_needed"], r["domains_needed"], max(1, r["domain_capacity"] // 100)],
            [CLR_PURPLE, "#B79AF3", CLR_INDIGO],
            "Email Infrastructure"), width='stretch',
            config={"displayModeBar": False}, key="b_email")

    # ══ ROTATION TIMELINE ═════════════════════════════════════════════════════
    if start_date and r.get("rotate_out_1"):
        st.divider()
        st.markdown("##### 🗓️&nbsp;&nbsp;**Rotation Schedule**")
        events = [
            {"task": "📞 Caller IDs", "start": start_date, "end": r["rotate_out_1"], "color": CLR_BLUE},
            {"task": "📞 Caller IDs (cyc 2)", "start": r["rotate_out_1"], "end": r["rotate_out_2"], "color": "#9AB6F5"},
            {"task": "💬 SMS Numbers", "start": start_date, "end": r["rotate_out_1"], "color": CLR_GREEN},
            {"task": "💬 SMS Numbers (cyc 2)", "start": r["rotate_out_1"], "end": r["rotate_out_2"], "color": "#7DD3A8"},
            {"task": "📧 Email Inboxes", "start": start_date, "end": r["email_rotate"], "color": CLR_PURPLE},
        ]
        if r.get("warmup_start"):
            events.insert(0, {"task": "📧 Email Warmup", "start": r["warmup_start"], "end": start_date, "color": CLR_AMBER})
        st.plotly_chart(timeline_chart(events), width='stretch',
                        config={"displayModeBar": False}, key="tl_rot")
        tcols = st.columns(3)
        kpi_card(tcols[0], "Campaign End", r["campaign_end"].strftime("%b %d, %Y"), "20 business days")
        kpi_card(tcols[1], "Warmup Start", r["warmup_start"].strftime("%b %d, %Y"), "Start − 30 days")
        kpi_card(tcols[2], "Email Rotation", r["email_rotate"].strftime("%b %d, %Y"), "After 20 sending days")

def page_reports():
    auto_phones = st.session_state.get("ppl_orig_phones", 0)
    auto_emails = st.session_state.get("ppl_orig_emails", 0)
    mkt_counts  = st.session_state.get("ppl_mkt_counts", {})
    auto_sms    = mkt_counts.get("sms", auto_phones)
    if mkt_counts.get("dialer"):
        auto_phones = mkt_counts["dialer"]
    if mkt_counts.get("email"):
        auto_emails = mkt_counts["email"]
    render_reporting(auto_phones=auto_phones, auto_sms=auto_sms, auto_emails=auto_emails)

# ── APP ──────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Campaign Organizer", page_icon="📋", layout="centered")

CUSTOM_CSS = """
<style>
#MainMenu, footer, header {visibility: hidden;}
html, body, [class*="css"] { font-family: 'Segoe UI', 'Inter', sans-serif; }
.stApp { background: #F4F6FB; }
[data-testid="stSidebarCollapseButton"] { display: none !important; }
section[data-testid="stSidebar"] .stButton>button {
    width: 100%; text-align: left; justify-content: flex-start;
    background: transparent !important; color: #4B5170 !important;
    border: 1px solid transparent !important; border-radius: 10px !important;
    font-weight: 600 !important; padding: 10px 14px !important;
    box-shadow: none !important; margin-bottom: 4px;
}
section[data-testid="stSidebar"] .stButton>button:hover {
    background: #F1F3FC !important; color: #3346D3 !important;
}
section[data-testid="stSidebar"] .stButton>button[kind="primary"] {
    background: #3346D3 !important; color: #FFFFFF !important;
    box-shadow: 0 4px 10px rgba(51,70,211,0.28) !important;
}
div[data-testid="stMain"] .stButton>button {
    border-radius: 10px !important; font-weight: 600 !important;
    padding: 10px 18px !important; border: 1px solid #D8DCEC !important;
}
div[data-testid="stMain"] .stButton>button[kind="primary"] {
    background: #3346D3 !important; color: #fff !important; border: none !important;
    box-shadow: 0 4px 10px rgba(51,70,211,0.25) !important;
}
.stDownloadButton>button {
    border-radius: 10px !important; font-weight: 600 !important;
    background: #1F9D55 !important; color: #fff !important; border: none !important;
    box-shadow: 0 4px 10px rgba(31,157,85,0.2) !important;
}
div[data-testid="stFileUploaderDropzone"] {
    border-radius: 12px !important; background: #FAFBFF !important;
    border: 2px dashed #C3C9E6 !important;
}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

NAV_PAGES = [("People Leads","👤"), ("Business Leads","🏢"), ("Regarding Reports","📊")]

if "current_page" not in st.session_state:
    st.session_state.current_page = "People Leads"

with st.sidebar:
    st.markdown("## 📋 Campaign Organizer")
    st.divider()
    for name, icon in NAV_PAGES:
        is_active = st.session_state.current_page == name
        if st.button(f"{icon}  {name}", key=f"nav_{name}",
                     type="primary" if is_active else "secondary",
                     width='stretch'):
            st.session_state.current_page = name
            st.rerun()

page = st.session_state.current_page
if page == "People Leads":
    page_people_leads()
elif page == "Business Leads":
    page_business_leads()
else:
    page_reports()
