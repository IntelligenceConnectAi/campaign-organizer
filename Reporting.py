import streamlit as st
import math
from datetime import date, timedelta, datetime
import pandas as pd

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
