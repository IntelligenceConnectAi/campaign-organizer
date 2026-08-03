import streamlit as st
import pandas as pd
import io
import math
import zipfile
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

# ── SHARED HELPERS ───────────────────────────────────────────────────────────
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

def merge_uploaded(files):
    frames = []
    orig_phones = 0
    orig_emails = 0
    phone_cols = [p for p, t, e in PHONE_GROUPS]
    email_cols = [e for p, t, e in PHONE_GROUPS]
    for uf in files:
        df = pd.read_excel(uf, dtype=str)
        df.drop(columns=[c for c in DNC_COLS if c in df.columns], inplace=True)
        for col in phone_cols:
            if col in df.columns:
                orig_phones += df[col].apply(
                    lambda x: 1 if pd.notna(x) and str(x).strip() not in ("", "nan") else 0).sum()
        for col in email_cols:
            if col in df.columns:
                orig_emails += df[col].apply(
                    lambda x: 1 if pd.notna(x) and str(x).strip() not in ("", "nan") else 0).sum()
        frames.append(df)
    merged = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    return merged, int(orig_phones), int(orig_emails)

def filter_properties(df):
    if "MLS Status" not in df.columns:
        return pd.DataFrame(columns=df.columns)
    blank_mask = df["MLS Status"].apply(lambda x: str(x).strip() in ("", "nan"))
    known_mask = df["MLS Status"].apply(lambda x: str(x).strip().upper() in MLS_STATUSES)
    known  = df[known_mask].copy()
    blanks = df[blank_mask].copy()
    blanks["MLS Status"] = "Off Market"
    return pd.concat([known, blanks], ignore_index=True)

def process_dialer_df(df, campaign_name):
    all_phone_cols = [p for p, t, e in PHONE_GROUPS] + [t for p, t, e in PHONE_GROUPS]
    other_cols     = [c for c in df.columns if c not in all_phone_cols]
    active_groups  = [(p, t) for p, t, e in PHONE_GROUPS if p in df.columns]
    output_cols    = ["Campaign Name", "Phone Number", "Phone Type"] + other_cols
    rows = []
    for _, row in df.iterrows():
        for phone_col, type_col in active_groups:
            phone_val = clean_phone(get_val(row, phone_col))
            if not phone_val:
                continue
            new_row = {"Campaign Name": campaign_name, "Phone Number": phone_val, "Phone Type": get_val(row, type_col)}
            for col in other_cols:
                new_row[col] = get_val(row, col)
            rows.append(new_row)
    return pd.DataFrame(rows, columns=output_cols)

def process_sms_df(df, campaign_name):
    all_pe_cols = ([p for p, t, e in PHONE_GROUPS] + [t for p, t, e in PHONE_GROUPS] +
                   [e for p, t, e in PHONE_GROUPS if e in df.columns])
    other_cols    = [c for c in df.columns if c not in all_pe_cols]
    active_groups = [(p, t, e) for p, t, e in PHONE_GROUPS if p in df.columns or e in df.columns]
    output_cols   = ["Campaign Name", "Phone Number", "Phone Type", "Email"] + other_cols
    rows = []
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
            rows.append(new_row)
    return pd.DataFrame(rows, columns=output_cols)

def process_email_df(df, campaign_name):
    all_pe_cols = ([p for p, t, e in PHONE_GROUPS] + [t for p, t, e in PHONE_GROUPS] +
                   [e for p, t, e in PHONE_GROUPS if e in df.columns])
    other_cols    = [c for c in df.columns if c not in all_pe_cols]
    active_groups = [(p, t, e) for p, t, e in PHONE_GROUPS if p in df.columns or e in df.columns]
    output_cols   = ["Campaign Name", "Phone Number", "Phone Type", "Email"] + other_cols
    rows = []
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
            rows.append(new_row)
    return pd.DataFrame(rows, columns=output_cols)

def split_parts(df, name_builder, n_splits, do_split):
    """Returns list of (filename, csv_bytes). name_builder(wk_or_none) -> campaign name string."""
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
            buf = io.StringIO()
            c.to_csv(buf, index=False)
            parts.append((name, buf.getvalue().encode("utf-8")))
        return parts
    else:
        name = name_builder(None)
        buf = io.StringIO()
        df.to_csv(buf, index=False)
        return [(name, buf.getvalue().encode("utf-8"))]

def build_zip(channel_parts):
    """channel_parts: dict of folder_name -> list[(filename, bytes)]"""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for folder, parts in channel_parts.items():
            for name, data in parts:
                zf.writestr(f"{folder}/{name}.csv", data)
    buf.seek(0)
    return buf.getvalue()

def step_header(num, icon, title, optional=False):
    badge = f'<span class="badge-optional">optional</span>' if optional else ""
    st.markdown(
        f'''<div class="step-header">
                <span class="step-num">{num}</span>
                <span class="step-icon">{icon}</span>
                <span class="step-title">{title}</span>
                {badge}
            </div>''',
        unsafe_allow_html=True,
    )

def page_title(icon, title, subtitle):
    st.markdown(
        f'''<div class="page-hero">
                <div class="page-hero-icon">{icon}</div>
                <div>
                    <div class="page-hero-title">{title}</div>
                    <div class="page-hero-sub">{subtitle}</div>
                </div>
            </div>''',
        unsafe_allow_html=True,
    )

def campaign_details_block(key_prefix):
    """Renders the 'Campaign Name' toggle + fields. Returns (add_name, month, year, state, deal, output_name)."""
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

def marketing_process_block(key_prefix, label_prefix=""):
    col1, col2, col3 = st.columns(3)
    with col1:
        dialer = st.checkbox("📞 Dialer / AI Outbound", key=f"{key_prefix}_dialer")
    with col2:
        sms = st.checkbox("💬 SMS", key=f"{key_prefix}_sms")
    with col3:
        email = st.checkbox("📧 Email", key=f"{key_prefix}_email")
    return dialer, sms, email

def run_channel_processing(df, tail, tag, dialer, sms, email, do_split, n_splits):
    """Runs dialer/sms/email processing for a given source df + campaign tag. Returns dict channel_parts + row counts."""
    channel_parts = {}
    counts = {}
    if dialer:
        out_df = process_dialer_df(df, build_campaign("CC", tail, tag=tag))
        counts["dialer"] = len(out_df)
        channel_parts["CC"] = split_parts(out_df, lambda wk: build_campaign("CC", tail, tag=tag, wk=wk), n_splits, do_split)
    if sms:
        out_df = process_sms_df(df, build_campaign("SMS", tail, tag=tag))
        counts["sms"] = len(out_df)
        channel_parts["SMS"] = split_parts(out_df, lambda wk: build_campaign("SMS", tail, tag=tag, wk=wk), n_splits, do_split)
    if email:
        out_df = process_email_df(df, build_campaign("EMAIL", tail, tag=tag))
        counts["email"] = len(out_df)
        channel_parts["EMAIL"] = split_parts(out_df, lambda wk: build_campaign("EMAIL", tail, tag=tag, wk=wk), n_splits, do_split)
    return channel_parts, counts

# ── PAGE: PEOPLE LEADS ──────────────────────────────────────────────────────
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
        n_splits = st.number_input("How many files to split into?", min_value=2, max_value=50, value=5, step=1, key="ppl_nsplits")

    mkt_selected  = dialer or sms or email
    prop_selected = p_dialer or p_sms or p_email
    details_ok = (out_name.strip() if not add_name and out_name is not None else True) and \
                 (state and year and deal and state.strip() and year.strip() and deal.strip() if add_name else (out_name and out_name.strip()))
    ready = files and details_ok and (mkt_selected or prop_selected)

    if st.button("⚙️ Process Files", use_container_width=True, type="primary", disabled=not ready, key="ppl_process"):
        merged, orig_phones, orig_emails = merge_uploaded(files)
        tail = get_tail(add_name, month, year, state, deal, out_name)

        mkt_parts, mkt_counts = run_channel_processing(merged, tail, None, dialer, sms, email, do_split, int(n_splits))
        st.session_state.ppl_mkt_zip = build_zip(mkt_parts) if mkt_parts else None
        st.session_state.ppl_mkt_name = f"Marketing Process - {tail}.zip"
        st.session_state.ppl_mkt_counts = mkt_counts

        prop_parts, prop_counts = ({}, {})
        if prop_selected:
            props_df = filter_properties(merged)
            prop_parts, prop_counts = run_channel_processing(props_df, tail, "PROPERTIES", p_dialer, p_sms, p_email, do_split, int(n_splits))
        st.session_state.ppl_prop_zip = build_zip(prop_parts) if prop_parts else None
        st.session_state.ppl_prop_name = f"Properties Seller Leads - {tail}.zip"
        st.session_state.ppl_prop_counts = prop_counts

        st.session_state.ppl_total = len(merged)
        st.session_state.ppl_orig_phones = orig_phones
        st.session_state.ppl_orig_emails = orig_emails
        st.session_state.ppl_processed = True

    if st.session_state.get("ppl_processed"):
        st.markdown('<div class="result-banner">✅ Processing Complete!</div>', unsafe_allow_html=True)
        st.info(f"📊 Total rows merged: **{st.session_state.ppl_total:,}**")
        st.info(f"📞 Original phone numbers: **{st.session_state.ppl_orig_phones:,}**")
        st.info(f"📧 Original emails: **{st.session_state.ppl_orig_emails:,}**")

        if st.session_state.get("ppl_mkt_counts"):
            st.markdown("**📋 Marketing Process:**")
            for k, v in st.session_state.ppl_mkt_counts.items():
                st.success(f"{k.title()}: **{v:,}**")
        if st.session_state.get("ppl_prop_counts"):
            st.markdown("**🏷️ Properties / Seller Leads:**")
            for k, v in st.session_state.ppl_prop_counts.items():
                st.success(f"{k.title()}: **{v:,}**")

        step_header("⬇", "📦", "Downloads")
        if st.session_state.get("ppl_mkt_zip"):
            st.download_button("⬇️ Download Marketing Process (ZIP)", data=st.session_state.ppl_mkt_zip,
                                file_name=st.session_state.ppl_mkt_name, mime="application/zip",
                                use_container_width=True, type="primary", key="ppl_dl_mkt")
        if st.session_state.get("ppl_prop_zip"):
            st.download_button("⬇️ Download Properties / Seller Leads (ZIP)", data=st.session_state.ppl_prop_zip,
                                file_name=st.session_state.ppl_prop_name, mime="application/zip",
                                use_container_width=True, type="primary", key="ppl_dl_prop")

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
            st.warning("⚠️ Please select at least one output type (Marketing Process or Properties).")

# ── PAGE: BUSINESS LEADS ────────────────────────────────────────────────────
def simple_merge_uploaded(files):
    """Business Leads have a different column schema than People Leads —
    no DNC-column drop, no phone/email normalization. Just merge as-is."""
    frames = []
    for uf in files:
        df = pd.read_excel(uf, dtype=str)
        frames.append(df)
    merged = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    return merged

def page_business_leads():
    page_title("🏢", "Business Leads", "Merge your business lead files and export in one click")

    step_header(1, "📁", "Upload Lead List")
    files = st.file_uploader("Upload one or more Excel files (.xlsx)",
                              type=["xlsx"], accept_multiple_files=True, key="biz_upload",
                              label_visibility="collapsed")
    if files:
        st.success(f"✅ {len(files)} file(s) uploaded")

    step_header(2, "🏷️", "Campaign Name")
    add_name, month, year, state, deal, out_name = campaign_details_block("biz")

    step_header(3, "🔀", "Monthly Split", optional=True)
    do_split = st.checkbox("🔀 Split output into multiple files", key="biz_dosplit")
    n_splits = 1
    if do_split:
        n_splits = st.number_input("How many files to split into?", min_value=2, max_value=50, value=5, step=1, key="biz_nsplits")

    details_ok = (state and year and deal and state.strip() and year.strip() and deal.strip()) if add_name else (out_name and out_name.strip())
    ready = files and details_ok

    if st.button("⚙️ Process Files", use_container_width=True, type="primary", disabled=not ready, key="biz_process"):
        merged = simple_merge_uploaded(files)
        tail = get_tail(add_name, month, year, state, deal, out_name)

        # Business Leads output — just the merged file, tagged with campaign name, optionally split
        base_name_builder = lambda wk: build_campaign("LEADS", tail, wk=wk)
        base_parts = split_parts(merged, base_name_builder, int(n_splits), do_split)
        st.session_state.biz_base_zip  = build_zip({"LEADS": base_parts}) if base_parts else None
        st.session_state.biz_base_name = f"Business Leads Output - {tail}.zip"
        st.session_state.biz_base_count = len(merged)

        st.session_state.biz_total = len(merged)
        st.session_state.biz_processed = True

    if st.session_state.get("biz_processed"):
        st.markdown('<div class="result-banner">✅ Processing Complete!</div>', unsafe_allow_html=True)
        st.info(f"📊 Total rows merged: **{st.session_state.biz_total:,}**")
        st.success(f"📄 Business Leads Output rows: **{st.session_state.biz_base_count:,}**")

        step_header("⬇", "📦", "Downloads")
        if st.session_state.get("biz_base_zip"):
            st.download_button("⬇️ Download Business Leads Output (ZIP)", data=st.session_state.biz_base_zip,
                                file_name=st.session_state.biz_base_name, mime="application/zip",
                                use_container_width=True, type="primary", key="biz_dl_base")

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

# ── PAGE: REGARDING REPORTS ─────────────────────────────────────────────────
def page_reports():
    page_title("📊", "Regarding Reports", "Campaign performance reporting — coming soon")
    st.markdown(
        '''<div class="empty-state">
                <div class="empty-state-icon">🚧</div>
                <div class="empty-state-title">Coming Soon</div>
                <div class="empty-state-sub">This section will be built out in a future update.</div>
            </div>''',
        unsafe_allow_html=True,
    )

# ── APP ──────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Campaign Organizer", page_icon="📋", layout="centered")

CUSTOM_CSS = """
<style>
#MainMenu, footer, header {visibility: hidden;}

html, body, [class*="css"] { font-family: 'Segoe UI', 'Inter', sans-serif; }

.stApp { background: linear-gradient(180deg, #f7f9fc 0%, #eef1f8 100%); }

section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #1b2340 0%, #262f52 100%);
}
section[data-testid="stSidebar"] * { color: #e7eaf6 !important; }
section[data-testid="stSidebar"] .sidebar-brand {
    display: flex; align-items: center; gap: 10px;
    padding: 22px 4px 18px 4px; margin-bottom: 6px;
    border-bottom: 1px solid rgba(255,255,255,0.12);
}
section[data-testid="stSidebar"] .sidebar-brand-icon {
    font-size: 28px;
}
section[data-testid="stSidebar"] .sidebar-brand-text {
    font-size: 18px; font-weight: 700; line-height: 1.15;
}
section[data-testid="stSidebar"] .sidebar-brand-sub {
    font-size: 11px; opacity: 0.65; font-weight: 400;
}

.page-hero {
    display: flex; align-items: center; gap: 16px;
    background: linear-gradient(135deg, #2f3a6b 0%, #4a5aa8 100%);
    border-radius: 16px; padding: 22px 26px; margin-bottom: 26px;
    box-shadow: 0 6px 20px rgba(47,58,107,0.25);
}
.page-hero-icon {
    font-size: 34px; background: rgba(255,255,255,0.15);
    border-radius: 12px; width: 56px; height: 56px;
    display: flex; align-items: center; justify-content: center;
}
.page-hero-title { color: #ffffff; font-size: 24px; font-weight: 700; }
.page-hero-sub { color: #dfe3f7; font-size: 13px; margin-top: 2px; }

.step-header {
    display: flex; align-items: center; gap: 10px;
    margin: 26px 0 10px 0; padding-bottom: 8px;
    border-bottom: 2px solid #e3e7f3;
}
.step-num {
    background: #4a5aa8; color: #fff; font-weight: 700; font-size: 13px;
    width: 24px; height: 24px; border-radius: 50%;
    display: flex; align-items: center; justify-content: center; flex-shrink: 0;
}
.step-icon { font-size: 18px; }
.step-title { font-size: 16px; font-weight: 700; color: #1f2748; }
.badge-optional {
    margin-left: auto; font-size: 10px; font-weight: 700; letter-spacing: 0.04em;
    text-transform: uppercase; color: #8a5a00; background: #ffe8b3;
    padding: 3px 9px; border-radius: 20px;
}

.result-banner {
    background: linear-gradient(135deg, #1f9d55 0%, #2fbf71 100%);
    color: #fff; font-weight: 700; font-size: 16px;
    padding: 14px 20px; border-radius: 12px; margin: 22px 0 14px 0;
    box-shadow: 0 4px 14px rgba(31,157,85,0.25);
}

.empty-state {
    text-align: center; padding: 60px 20px;
    background: #ffffff; border-radius: 16px;
    border: 1px dashed #c9cee0;
}
.empty-state-icon { font-size: 42px; margin-bottom: 10px; }
.empty-state-title { font-size: 18px; font-weight: 700; color: #1f2748; }
.empty-state-sub { font-size: 13px; color: #6b7290; margin-top: 4px; }

.stButton>button {
    border-radius: 10px !important; font-weight: 600 !important;
    padding: 10px 18px !important; border: none !important;
}
.stButton>button[kind="primary"] {
    background: linear-gradient(135deg, #4a5aa8 0%, #2f3a6b 100%) !important;
    box-shadow: 0 4px 12px rgba(47,58,107,0.3) !important;
}
.stDownloadButton>button {
    border-radius: 10px !important; font-weight: 600 !important;
    background: linear-gradient(135deg, #1f9d55 0%, #2fbf71 100%) !important;
    color: #fff !important; border: none !important;
    box-shadow: 0 4px 12px rgba(31,157,85,0.25) !important;
}

div[data-testid="stFileUploaderDropzone"] {
    border-radius: 14px !important; background: #ffffff !important;
    border: 2px dashed #b7c0e0 !important;
}

.stCheckbox, .stTextInput, .stSelectbox, .stNumberInput { margin-bottom: 2px; }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

with st.sidebar:
    st.markdown(
        '''<div class="sidebar-brand">
                <div class="sidebar-brand-icon">📋</div>
                <div>
                    <div class="sidebar-brand-text">Campaign<br/>Organizer</div>
                </div>
            </div>''',
        unsafe_allow_html=True,
    )
    try:
        from streamlit_option_menu import option_menu
        page = option_menu(
            menu_title=None,
            options=["People Leads", "Business Leads", "Regarding Reports"],
            icons=["person-lines-fill", "briefcase-fill", "bar-chart-line-fill"],
            default_index=0,
            styles={
                "container": {"padding": "0!important", "background-color": "transparent"},
                "icon": {"color": "#9aa4d6", "font-size": "16px"},
                "nav-link": {
                    "font-size": "14px", "font-weight": "600", "text-align": "left",
                    "margin": "4px 0", "border-radius": "10px", "color": "#c8cde8",
                    "padding": "11px 14px",
                },
                "nav-link-selected": {
                    "background": "linear-gradient(135deg, #4a5aa8 0%, #2f3a6b 100%)",
                    "color": "#ffffff",
                },
            },
        )
    except ImportError:
        st.caption("Tip: `pip install streamlit-option-menu` for a nicer nav.")
        page = st.radio(
            "Navigation",
            ["People Leads", "Business Leads", "Regarding Reports"],
            key="nav_page", label_visibility="collapsed",
        )

if page == "People Leads":
    page_people_leads()
elif page == "Business Leads":
    page_business_leads()
else:
    page_reports()
