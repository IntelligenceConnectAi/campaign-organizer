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

    if st.button("⚙️ Process Files", use_container_width=True, type="primary",
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
                               mime="application/zip", use_container_width=True,
                               type="primary", key="ppl_dl_mkt")
        if st.session_state.get("ppl_prop_zip"):
            st.download_button("⬇️ Download Properties / Seller Leads (ZIP)",
                               data=st.session_state.ppl_prop_zip,
                               file_name=st.session_state.ppl_prop_name,
                               mime="application/zip", use_container_width=True,
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

    if st.button("⚙️ Process Files", use_container_width=True, type="primary",
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
                               mime="application/zip", use_container_width=True,
                               type="primary", key="biz_dl")

    if not ready and files:
        if add_name:
            if not (state and state.strip()): st.warning("⚠️ Please enter State.")
            if not (year and year.strip()):   st.warning("⚠️ Please enter Year.")
            if not (deal and deal.strip()):   st.warning("⚠️ Please enter Type of Deal.")
        else:
            if not (out_name and out_name.strip()): st.warning("⚠️ Please enter Output File Name.")

# ── PAGE: REPORTS ────────────────────────────────────────────────────────────
# ── CONSTANTS FROM RULES TABS ─────────────────────────────────────────────────
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

    # ── COLD CALLING ──────────────────────────────────────────────────────────
    r["dials_per_day"]        = math.ceil(total_phones / CALLING_DAYS_PER_MONTH)
    r["callers_needed"]       = math.ceil(r["dials_per_day"] / DIALS_PER_CALLER)
    r["active_caller_ids"]    = max(MIN_CALLER_ID_POOL, math.ceil(r["dials_per_day"] / MAX_DIALS_PER_NUMBER))
    r["rotation_cycles"]      = CALLING_DAYS_PER_MONTH // ROTATION_CALLING_DAYS
    r["total_caller_ids"]     = r["active_caller_ids"] * r["rotation_cycles"]
    r["dials_per_number"]     = math.ceil(r["dials_per_day"] / r["active_caller_ids"])
    r["dials_check"]          = "✅ OK" if r["dials_per_number"] <= MAX_DIALS_PER_NUMBER else "⛔ OVER LIMIT"
    r["dials_check_ok"]       = r["dials_per_number"] <= MAX_DIALS_PER_NUMBER

    # ── STAFFING ─────────────────────────────────────────────────────────────
    r["callers_diff"]         = callers - r["callers_needed"]
    r["staffing_ok"]          = callers >= r["callers_needed"]
    r["team_capacity"]        = callers * DIALS_PER_CALLER * CALLING_DAYS_PER_MONTH
    r["days_to_finish"]       = math.ceil(total_phones / (callers * DIALS_PER_CALLER)) if callers > 0 else 999

    # ── SMS ──────────────────────────────────────────────────────────────────
    r["texts_per_day"]        = math.ceil(total_sms / CALLING_DAYS_PER_MONTH)
    r["min_sms_numbers"]      = math.ceil(r["texts_per_day"] / MAX_SMS_PER_NUMBER)
    r["campaigns_required"]   = math.ceil(r["texts_per_day"] / MAX_SMS_PER_CAMPAIGN)
    r["active_sms_numbers"]   = max(r["min_sms_numbers"], r["campaigns_required"] * SMS_NUMBERS_PER_CAMPAIGN)
    r["total_sms_numbers"]    = r["active_sms_numbers"] * r["rotation_cycles"]
    r["texts_per_number"]     = math.ceil(r["texts_per_day"] / r["active_sms_numbers"]) if r["active_sms_numbers"] > 0 else 0
    r["sms_check"]            = "✅ OK" if r["texts_per_number"] <= MAX_SMS_PER_NUMBER else "⛔ OVER LIMIT"
    r["sms_check_ok"]         = r["texts_per_number"] <= MAX_SMS_PER_NUMBER

    # ── EMAIL ─────────────────────────────────────────────────────────────────
    r["emails_per_day"]       = math.ceil(total_emails / CALLING_DAYS_PER_MONTH)
    r["inboxes_needed"]       = math.ceil(r["emails_per_day"] / MAX_EMAILS_PER_INBOX)
    r["domains_needed"]       = math.ceil(r["inboxes_needed"] / MAX_INBOXES_PER_DOMAIN)
    r["domain_capacity"]      = r["domains_needed"] * DOMAIN_DAILY_CAPACITY
    r["emails_per_inbox"]     = math.ceil(r["emails_per_day"] / r["inboxes_needed"]) if r["inboxes_needed"] > 0 else 0
    r["email_check"]          = "✅ OK" if r["emails_per_inbox"] <= MAX_EMAILS_PER_INBOX else "⛔ OVER LIMIT"
    r["email_check_ok"]       = r["emails_per_inbox"] <= MAX_EMAILS_PER_INBOX

    # ── DATES ─────────────────────────────────────────────────────────────────
    if start_date:
        r["campaign_end"]       = add_business_days(start_date, CALLING_DAYS_PER_MONTH)
        r["rotate_out_1"]       = add_business_days(start_date, ROTATION_CALLING_DAYS)
        r["rotate_out_2"]       = add_business_days(start_date, ROTATION_CALLING_DAYS * 2)
        r["warmup_start"]       = start_date - timedelta(days=EMAIL_WARMUP_DAYS)
        r["email_rotate"]       = add_business_days(start_date, EMAIL_ROTATION_DAYS)
    else:
        r["campaign_end"]       = None
        r["rotate_out_1"]       = None
        r["rotate_out_2"]       = None
        r["warmup_start"]       = None
        r["email_rotate"]       = None

    return r


def render_reporting(auto_phones=0, auto_sms=0, auto_emails=0):

    # ── CSS ───────────────────────────────────────────────────────────────────
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');

    .report-wrap { font-family: 'Inter', sans-serif; }

    .hero-title {
        text-align: center;
        font-size: 2.2rem;
        font-weight: 900;
        letter-spacing: -0.5px;
        background: linear-gradient(135deg, #F5A623 0%, #F7C948 50%, #F5A623 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        margin-bottom: 4px;
    }
    .hero-sub {
        text-align: center;
        color: #8892A4;
        font-size: 0.9rem;
        font-weight: 400;
        letter-spacing: 2px;
        text-transform: uppercase;
        margin-bottom: 32px;
    }
    .section-title {
        font-size: 0.72rem;
        font-weight: 700;
        letter-spacing: 3px;
        text-transform: uppercase;
        color: #F5A623;
        margin-bottom: 16px;
        margin-top: 8px;
    }

    /* ── METRIC CARD ─────────────────────────────────── */
    .metric-card {
        background: linear-gradient(145deg, #1A2035 0%, #141929 100%);
        border: 1px solid rgba(245,166,35,0.15);
        border-radius: 16px;
        padding: 20px 22px;
        margin-bottom: 14px;
        box-shadow: 0 4px 24px rgba(0,0,0,0.3);
        transition: all 0.2s;
    }
    .metric-card:hover {
        border-color: rgba(245,166,35,0.35);
        box-shadow: 0 8px 32px rgba(245,166,35,0.08);
        transform: translateY(-1px);
    }
    .metric-label {
        font-size: 0.72rem;
        font-weight: 600;
        color: #8892A4;
        text-transform: uppercase;
        letter-spacing: 1.5px;
        margin-bottom: 6px;
    }
    .metric-value {
        font-size: 2.1rem;
        font-weight: 800;
        color: #F5A623;
        line-height: 1;
        margin-bottom: 4px;
    }
    .metric-desc {
        font-size: 0.72rem;
        color: #5A6378;
        font-weight: 400;
    }
    .metric-value-white {
        font-size: 2.1rem;
        font-weight: 800;
        color: #E8EAF0;
        line-height: 1;
        margin-bottom: 4px;
    }

    /* ── CHANNEL HEADER ─────────────────────────────── */
    .channel-header {
        border-radius: 14px;
        padding: 18px 22px;
        margin-bottom: 20px;
        display: flex;
        align-items: center;
        gap: 12px;
    }
    .ch-cc   { background: linear-gradient(135deg, #1E3A5F 0%, #152D4A 100%); border: 1px solid rgba(59,130,246,0.3); }
    .ch-sms  { background: linear-gradient(135deg, #1E4535 0%, #153324 100%); border: 1px solid rgba(34,197,94,0.3); }
    .ch-mail { background: linear-gradient(135deg, #3D1E5F 0%, #2D1548 100%); border: 1px solid rgba(168,85,247,0.3); }
    .ch-icon { font-size: 2rem; }
    .ch-title { font-size: 1.1rem; font-weight: 800; color: #E8EAF0; }
    .ch-sub   { font-size: 0.75rem; color: #8892A4; font-weight: 400; }

    /* ── STATUS BADGES ──────────────────────────────── */
    .badge-ok     { background: rgba(34,197,94,0.15); color: #22C55E; border: 1px solid rgba(34,197,94,0.3); border-radius: 8px; padding: 4px 12px; font-size: 0.75rem; font-weight: 700; display: inline-block; }
    .badge-warn   { background: rgba(245,166,35,0.15); color: #F5A623; border: 1px solid rgba(245,166,35,0.3); border-radius: 8px; padding: 4px 12px; font-size: 0.75rem; font-weight: 700; display: inline-block; }
    .badge-danger { background: rgba(239,68,68,0.15); color: #EF4444; border: 1px solid rgba(239,68,68,0.3); border-radius: 8px; padding: 4px 12px; font-size: 0.75rem; font-weight: 700; display: inline-block; }

    /* ── STAFFING ALERT ─────────────────────────────── */
    .alert-danger {
        background: linear-gradient(135deg, rgba(239,68,68,0.12) 0%, rgba(185,28,28,0.08) 100%);
        border: 1px solid rgba(239,68,68,0.4);
        border-left: 4px solid #EF4444;
        border-radius: 14px;
        padding: 20px 24px;
        margin: 8px 0 20px 0;
    }
    .alert-success {
        background: linear-gradient(135deg, rgba(34,197,94,0.1) 0%, rgba(21,128,61,0.06) 100%);
        border: 1px solid rgba(34,197,94,0.35);
        border-left: 4px solid #22C55E;
        border-radius: 14px;
        padding: 20px 24px;
        margin: 8px 0 20px 0;
    }
    .alert-title  { font-size: 1.1rem; font-weight: 800; margin-bottom: 6px; }
    .alert-body   { font-size: 0.85rem; color: #8892A4; }

    /* ── ROTATION TABLE ─────────────────────────────── */
    .rot-card {
        background: linear-gradient(145deg, #1A2035 0%, #141929 100%);
        border: 1px solid rgba(245,166,35,0.12);
        border-radius: 16px;
        padding: 20px 24px;
        margin-bottom: 12px;
    }
    .rot-asset { font-size: 0.85rem; font-weight: 700; color: #E8EAF0; margin-bottom: 6px; }
    .rot-row   { display: flex; gap: 20px; flex-wrap: wrap; }
    .rot-item  { }
    .rot-key   { font-size: 0.65rem; color: #5A6378; text-transform: uppercase; letter-spacing: 1px; font-weight: 600; }
    .rot-val   { font-size: 0.9rem; color: #F5A623; font-weight: 700; }

    /* ── DIVIDER ────────────────────────────────────── */
    .gold-divider {
        height: 1px;
        background: linear-gradient(90deg, transparent, rgba(245,166,35,0.4), transparent);
        margin: 28px 0;
    }

    /* ── INPUT PANEL ────────────────────────────────── */
    .input-panel {
        background: linear-gradient(145deg, #1A2035 0%, #141929 100%);
        border: 1px solid rgba(245,166,35,0.2);
        border-radius: 18px;
        padding: 24px;
        margin-bottom: 28px;
    }
    </style>
    """, unsafe_allow_html=True)

    st.markdown('<div class="report-wrap">', unsafe_allow_html=True)

    # ── HERO ──────────────────────────────────────────────────────────────────
    st.markdown('<div class="hero-title">⚡ FHO MARKETING GUARDRAILS</div>', unsafe_allow_html=True)
    st.markdown('<div class="hero-sub">Monthly Campaign Intelligence Dashboard</div>', unsafe_allow_html=True)

    # ── INPUT PANEL ───────────────────────────────────────────────────────────
    st.markdown('<div class="section-title">▸ Campaign Inputs</div>', unsafe_allow_html=True)
    with st.container():
        c1, c2, c3 = st.columns(3)
        with c1:
            total_phones = st.number_input("📞 Total Phone Numbers",
                min_value=0, value=int(auto_phones) if auto_phones else 0,
                step=1000, format="%d", key="rpt_phones")
        with c2:
            total_sms = st.number_input("💬 Total SMS Numbers (Mobile)",
                min_value=0, value=int(auto_sms) if auto_sms else 0,
                step=1000, format="%d", key="rpt_sms")
        with c3:
            total_emails = st.number_input("📧 Total Email Addresses",
                min_value=0, value=int(auto_emails) if auto_emails else 0,
                step=1000, format="%d", key="rpt_emails")

        c4, c5 = st.columns(2)
        with c4:
            callers = st.number_input("👥 Callers on Team",
                min_value=0, value=2, step=1, format="%d", key="rpt_callers")
        with c5:
            use_dates = st.checkbox("📅 Add Campaign Dates (Optional)", key="rpt_use_dates")

        start_date = None
        if use_dates:
            start_date = st.date_input("Campaign Start Date", value=date.today(), key="rpt_start")

    if total_phones == 0 and total_sms == 0 and total_emails == 0:
        st.markdown("""
        <div style="text-align:center; padding: 60px 20px; color: #5A6378;">
            <div style="font-size: 3rem; margin-bottom: 12px;">📊</div>
            <div style="font-size: 1rem; font-weight: 600; color: #8892A4;">Enter your numbers above to generate the report</div>
            <div style="font-size: 0.8rem; margin-top: 6px;">Process your files first — counts will auto-populate</div>
        </div>
        """, unsafe_allow_html=True)
        return

    r = calculate_all(total_phones, total_sms, total_emails, callers, start_date)

    st.markdown('<div class="gold-divider"></div>', unsafe_allow_html=True)

    # ── COLD CALLING ─────────────────────────────────────────────────────────
    st.markdown("""
    <div class="channel-header ch-cc">
        <span class="ch-icon">📞</span>
        <div>
            <div class="ch-title">Cold Calling</div>
            <div class="ch-sub">Landline + Mobile · All phone numbers</div>
        </div>
    </div>""", unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-label">Dials per Day</div>
            <div class="metric-value">{r['dials_per_day']:,}</div>
            <div class="metric-desc">Total ÷ 20 days</div>
        </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-label">Callers Needed</div>
            <div class="metric-value">{r['callers_needed']:,}</div>
            <div class="metric-desc">Dials ÷ 600/caller</div>
        </div>""", unsafe_allow_html=True)
    with c3:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-label">Active Caller IDs</div>
            <div class="metric-value">{r['active_caller_ids']:,}</div>
            <div class="metric-desc">Per rotation cycle</div>
        </div>""", unsafe_allow_html=True)
    with c4:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-label">Total Caller IDs</div>
            <div class="metric-value">{r['total_caller_ids']:,}</div>
            <div class="metric-desc">Active × {r['rotation_cycles']} cycles</div>
        </div>""", unsafe_allow_html=True)

    c5, c6, c7, c8 = st.columns(4)
    with c5:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-label">Dials / Number / Day</div>
            <div class="metric-value {'metric-value' if r['dials_check_ok'] else 'metric-value'}" style="color: {'#22C55E' if r['dials_check_ok'] else '#EF4444'}">{r['dials_per_number']:,}</div>
            <div class="metric-desc">Limit: 150/day</div>
        </div>""", unsafe_allow_html=True)
    with c6:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-label">Rotation Cycles</div>
            <div class="metric-value-white">{r['rotation_cycles']}</div>
            <div class="metric-desc">Every 10 calling days</div>
        </div>""", unsafe_allow_html=True)
    with c7:
        color = "#22C55E" if r['dials_check_ok'] else "#EF4444"
        badge = "badge-ok" if r['dials_check_ok'] else "badge-danger"
        label = "WITHIN LIMIT" if r['dials_check_ok'] else "OVER LIMIT"
        st.markdown(f"""<div class="metric-card">
            <div class="metric-label">Dials Check</div>
            <div style="margin: 10px 0;"><span class="{badge}">{label}</span></div>
            <div class="metric-desc">Must stay ≤ 150/number/day</div>
        </div>""", unsafe_allow_html=True)
    with c8:
        if r.get("campaign_end"):
            st.markdown(f"""<div class="metric-card">
                <div class="metric-label">Campaign End</div>
                <div class="metric-value-white" style="font-size:1.3rem">{r['campaign_end'].strftime('%b %d, %Y')}</div>
                <div class="metric-desc">20 business days from start</div>
            </div>""", unsafe_allow_html=True)
        else:
            st.markdown(f"""<div class="metric-card">
                <div class="metric-label">Campaign Duration</div>
                <div class="metric-value-white">20</div>
                <div class="metric-desc">Business days</div>
            </div>""", unsafe_allow_html=True)

    st.markdown('<div class="gold-divider"></div>', unsafe_allow_html=True)

    # ── STAFFING ─────────────────────────────────────────────────────────────
    st.markdown('<div class="section-title">▸ Staffing Analysis</div>', unsafe_allow_html=True)

    if r["staffing_ok"]:
        surplus = r["callers_diff"]
        st.markdown(f"""<div class="alert-success">
            <div class="alert-title" style="color:#22C55E">✅ Staffing: Sufficient</div>
            <div class="alert-body">You have <strong style="color:#E8EAF0">{callers} callers</strong> — 
            {surplus} surplus over the {r['callers_needed']} needed. 
            Your team can finish the list in <strong style="color:#E8EAF0">{r['days_to_finish']} days</strong>.</div>
        </div>""", unsafe_allow_html=True)
    else:
        shortfall = abs(r["callers_diff"])
        st.markdown(f"""<div class="alert-danger">
            <div class="alert-title" style="color:#EF4444">⚠️ Staffing: Shortfall</div>
            <div class="alert-body">You have <strong style="color:#E8EAF0">{callers} callers</strong> but need 
            <strong style="color:#E8EAF0">{r['callers_needed']}</strong>. 
            Hire <strong style="color:#EF4444">{shortfall} more caller(s)</strong> to finish in 20 days. 
            At current staffing, the list takes <strong style="color:#EF4444">{r['days_to_finish']} days</strong>.</div>
        </div>""", unsafe_allow_html=True)

    sa, sb, sc = st.columns(3)
    with sa:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-label">Callers Needed</div>
            <div class="metric-value">{r['callers_needed']:,}</div>
            <div class="metric-desc">To finish in 20 days</div>
        </div>""", unsafe_allow_html=True)
    with sb:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-label">Team Capacity (20 days)</div>
            <div class="metric-value-white">{r['team_capacity']:,}</div>
            <div class="metric-desc">{callers} callers × 600 × 20</div>
        </div>""", unsafe_allow_html=True)
    with sc:
        days_color = "#22C55E" if r['days_to_finish'] <= 20 else "#EF4444"
        st.markdown(f"""<div class="metric-card">
            <div class="metric-label">Days to Finish</div>
            <div class="metric-value" style="color:{days_color}">{r['days_to_finish']}</div>
            <div class="metric-desc">Target: 20 or fewer</div>
        </div>""", unsafe_allow_html=True)

    st.markdown('<div class="gold-divider"></div>', unsafe_allow_html=True)

    # ── SMS ──────────────────────────────────────────────────────────────────
    st.markdown("""
    <div class="channel-header ch-sms">
        <span class="ch-icon">💬</span>
        <div>
            <div class="ch-title">SMS</div>
            <div class="ch-sub">Mobile numbers only · A2P 10DLC</div>
        </div>
    </div>""", unsafe_allow_html=True)

    s1, s2, s3, s4 = st.columns(4)
    with s1:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-label">Texts per Day</div>
            <div class="metric-value">{r['texts_per_day']:,}</div>
            <div class="metric-desc">SMS total ÷ 20 days</div>
        </div>""", unsafe_allow_html=True)
    with s2:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-label">Campaigns Required</div>
            <div class="metric-value">{r['campaigns_required']:,}</div>
            <div class="metric-desc">÷ 2,000/campaign/day</div>
        </div>""", unsafe_allow_html=True)
    with s3:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-label">Active SMS Numbers</div>
            <div class="metric-value">{r['active_sms_numbers']:,}</div>
            <div class="metric-desc">Per cycle · {SMS_NUMBERS_PER_CAMPAIGN}/campaign</div>
        </div>""", unsafe_allow_html=True)
    with s4:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-label">Total SMS Numbers</div>
            <div class="metric-value">{r['total_sms_numbers']:,}</div>
            <div class="metric-desc">Active × {r['rotation_cycles']} cycles</div>
        </div>""", unsafe_allow_html=True)

    s5, s6, s7, _ = st.columns(4)
    with s5:
        sms_color = "#22C55E" if r['sms_check_ok'] else "#EF4444"
        st.markdown(f"""<div class="metric-card">
            <div class="metric-label">Texts / Number / Day</div>
            <div class="metric-value" style="color:{sms_color}">{r['texts_per_number']:,}</div>
            <div class="metric-desc">Limit: 500/day</div>
        </div>""", unsafe_allow_html=True)
    with s6:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-label">Min Sending Numbers</div>
            <div class="metric-value-white">{r['min_sms_numbers']:,}</div>
            <div class="metric-desc">Texts ÷ 500/number</div>
        </div>""", unsafe_allow_html=True)
    with s7:
        badge = "badge-ok" if r['sms_check_ok'] else "badge-danger"
        label = "WITHIN LIMIT" if r['sms_check_ok'] else "OVER LIMIT"
        st.markdown(f"""<div class="metric-card">
            <div class="metric-label">SMS Check</div>
            <div style="margin: 10px 0;"><span class="{badge}">{label}</span></div>
            <div class="metric-desc">Must stay ≤ 500/number/day</div>
        </div>""", unsafe_allow_html=True)

    st.markdown('<div class="gold-divider"></div>', unsafe_allow_html=True)

    # ── EMAIL ─────────────────────────────────────────────────────────────────
    st.markdown("""
    <div class="channel-header ch-mail">
        <span class="ch-icon">📧</span>
        <div>
            <div class="ch-title">Email</div>
            <div class="ch-sub">Verified emails only · CAN-SPAM compliant</div>
        </div>
    </div>""", unsafe_allow_html=True)

    e1, e2, e3, e4 = st.columns(4)
    with e1:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-label">Emails per Day</div>
            <div class="metric-value">{r['emails_per_day']:,}</div>
            <div class="metric-desc">Total ÷ 20 days</div>
        </div>""", unsafe_allow_html=True)
    with e2:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-label">Inboxes Needed</div>
            <div class="metric-value">{r['inboxes_needed']:,}</div>
            <div class="metric-desc">÷ 35/inbox/day</div>
        </div>""", unsafe_allow_html=True)
    with e3:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-label">Domains Needed</div>
            <div class="metric-value">{r['domains_needed']:,}</div>
            <div class="metric-desc">÷ 5 inboxes/domain</div>
        </div>""", unsafe_allow_html=True)
    with e4:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-label">Domain Daily Capacity</div>
            <div class="metric-value-white">{r['domain_capacity']:,}</div>
            <div class="metric-desc">Domains × 175/day</div>
        </div>""", unsafe_allow_html=True)

    e5, e6, e7, e8 = st.columns(4)
    with e5:
        email_color = "#22C55E" if r['email_check_ok'] else "#EF4444"
        st.markdown(f"""<div class="metric-card">
            <div class="metric-label">Emails / Inbox / Day</div>
            <div class="metric-value" style="color:{email_color}">{r['emails_per_inbox']:,}</div>
            <div class="metric-desc">Limit: 35/day</div>
        </div>""", unsafe_allow_html=True)
    with e6:
        badge = "badge-ok" if r['email_check_ok'] else "badge-danger"
        label = "WITHIN LIMIT" if r['email_check_ok'] else "OVER LIMIT"
        st.markdown(f"""<div class="metric-card">
            <div class="metric-label">Email Check</div>
            <div style="margin: 10px 0;"><span class="{badge}">{label}</span></div>
            <div class="metric-desc">Must stay ≤ 35/inbox/day</div>
        </div>""", unsafe_allow_html=True)
    with e7:
        if r.get("warmup_start"):
            st.markdown(f"""<div class="metric-card">
                <div class="metric-label">Warmup Start Date</div>
                <div class="metric-value-white" style="font-size:1.2rem">{r['warmup_start'].strftime('%b %d, %Y')}</div>
                <div class="metric-desc">Start − 30 warmup days</div>
            </div>""", unsafe_allow_html=True)
        else:
            st.markdown(f"""<div class="metric-card">
                <div class="metric-label">Warmup Required</div>
                <div class="metric-value-white">30</div>
                <div class="metric-desc">Days before campaign</div>
            </div>""", unsafe_allow_html=True)
    with e8:
        if r.get("email_rotate"):
            st.markdown(f"""<div class="metric-card">
                <div class="metric-label">Inbox Rotation Date</div>
                <div class="metric-value-white" style="font-size:1.2rem">{r['email_rotate'].strftime('%b %d, %Y')}</div>
                <div class="metric-desc">After 20 sending days</div>
            </div>""", unsafe_allow_html=True)
        else:
            st.markdown(f"""<div class="metric-card">
                <div class="metric-label">Inbox Rotation</div>
                <div class="metric-value-white">20</div>
                <div class="metric-desc">Days per cycle</div>
            </div>""", unsafe_allow_html=True)

    # ── ROTATION SCHEDULE (only if dates provided) ────────────────────────────
    if start_date and r.get("rotate_out_1"):
        st.markdown('<div class="gold-divider"></div>', unsafe_allow_html=True)
        st.markdown('<div class="section-title">▸ Rotation Schedule</div>', unsafe_allow_html=True)

        rot_data = [
            {"Asset": "📞 Caller IDs", "Rotate Every": "10 calling days",
             "Trigger": "1,500 dials", "Rotate Out #1": r['rotate_out_1'].strftime('%b %d, %Y'),
             "Rotate Out #2": r['rotate_out_2'].strftime('%b %d, %Y') if r.get('rotate_out_2') else "—"},
            {"Asset": "💬 SMS Numbers", "Rotate Every": "10 sending days",
             "Trigger": "5,000 texts", "Rotate Out #1": r['rotate_out_1'].strftime('%b %d, %Y'),
             "Rotate Out #2": r['rotate_out_2'].strftime('%b %d, %Y') if r.get('rotate_out_2') else "—"},
            {"Asset": "📧 Email Inboxes", "Rotate Every": "20 sending days",
             "Trigger": "700 emails", "Rotate Out #1": r['email_rotate'].strftime('%b %d, %Y') if r.get('email_rotate') else "—",
             "Rotate Out #2": "—"},
        ]

        for row in rot_data:
            st.markdown(f"""
            <div class="rot-card">
                <div class="rot-asset">{row['Asset']}</div>
                <div class="rot-row">
                    <div class="rot-item"><div class="rot-key">Rotate Every</div><div class="rot-val">{row['Rotate Every']}</div></div>
                    <div class="rot-item"><div class="rot-key">Cumulative Trigger</div><div class="rot-val">{row['Trigger']}</div></div>
                    <div class="rot-item"><div class="rot-key">Rotate Out #1</div><div class="rot-val">{row['Rotate Out #1']}</div></div>
                    <div class="rot-item"><div class="rot-key">Rotate Out #2</div><div class="rot-val">{row['Rotate Out #2']}</div></div>
                </div>
            </div>""", unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)


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
                     use_container_width=True):
            st.session_state.current_page = name
            st.rerun()

page = st.session_state.current_page
if page == "People Leads":
    page_people_leads()
elif page == "Business Leads":
    page_business_leads()
else:
    page_reports()
